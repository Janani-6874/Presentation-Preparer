from __future__ import annotations
import io
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

os.system("pip install streamlit==1.38.0")
os.system("pip install python-pptx==0.6.23")
os.system("pip install requests>=2.31.0")

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.shapes import MSO_SHAPE_TYPE

import requests

APP_TITLE = "Your Text, Your Style – Auto-Generate a Presentation"
MAX_UPLOAD_MB = 20
MAX_TOKENS_FALLBACK = 8000  # naïve guardrail for giant inputs

# ------------------------------

# ------------------------------





def _write_repo_files_once():
    try:
        if not os.path.exists("LICENSE"):
            with open("LICENSE", "w", encoding="utf-8") as f:
                f.write(MIT_LICENSE_TEXT)
        if not os.path.exists("README.md"):
            with open("README.md", "w", encoding="utf-8") as f:
                f.write(README_MD)
        if not os.path.exists("requirements.txt"):
            with open("requirements.txt", "w", encoding="utf-8") as f:
                f.write(REQUIREMENTS_TXT)
    except Exception:
        # Non-fatal; don't surface in UI to avoid noisy errors
        pass


# ------------------------------
# LLM Provider Adapters
# ------------------------------
@dataclass
class LLMResponse:
    slides: List[Dict[str, Any]]


def call_openai(api_key: str, model: str, prompt: str, max_tokens: int = 2000) -> str:
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    body = {
        "model": model or "gpt-4o-mini",  # allow user override
        "messages": [
            {"role": "system", "content": "You are a helpful assistant that returns STRICT JSON only."},
            {"role": "user", "content": prompt},
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.2,
        "max_tokens": max_tokens,
    }
    r = requests.post(url, headers=headers, json=body, timeout=60)
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"].strip()


def call_anthropic(api_key: str, model: str, prompt: str, max_tokens: int = 2000) -> str:
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": model or "claude-3-5-sonnet-latest",
        "max_tokens": max_tokens,
        "temperature": 0.2,
        "messages": [
            {"role": "user", "content": prompt},
        ],
        "system": "Return STRICT JSON only.",
    }
    r = requests.post(url, headers=headers, json=body, timeout=60)
    r.raise_for_status()
    data = r.json()
    # Anthropic returns a list of content blocks; join text
    parts = [c.get("text", "") for c in data.get("content", []) if c.get("type") == "text"]
    return "".join(parts).strip()


def call_gemini(api_key: str, model: str, prompt: str, max_tokens: int = 2048) -> str:
    # Google Generative Language API (Gemini 1.5) – JSON may vary per account/region
    # Using the text endpoint with a JSON constraint in the prompt.
    model_name = model or "gemini-1.5-flash"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {"content-type": "application/json"}
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": max_tokens,
        },
    }
    r = requests.post(url, headers=headers, json=body, timeout=60)
    r.raise_for_status()
    data = r.json()
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        raise RuntimeError(f"Gemini response parse error: {e}\nRaw: {data}")


def outline_with_llm(provider: str, api_key: str, model: str, user_text: str, guidance: str, want_notes: bool) -> LLMResponse:
    schema_hint = {
        "slides": [
            {
                "title": "string",
                "bullets": ["string", "string"],
                **({"notes": "string"} if want_notes else {}),
            }
        ]
    }
    prompt = f"""
You are given INPUT_TEXT (markdown possible) and an optional GUIDANCE string describing tone/use-case.
Return a STRICT JSON object matching this schema (no prose):\n{json.dumps(schema_hint)}\n
Rules:\n- Choose a reasonable number of slides to cover the content succinctly.\n- Keep bullets concise (<=15 words each).\n- Preserve markdown structure (headings become slide sections when helpful).\n- {"Include a short 'notes' string for each slide." if want_notes else "Do NOT include 'notes'."}\n
GUIDANCE: {guidance or "(none)"}
INPUT_TEXT:\n\n{user_text[:MAX_TOKENS_FALLBACK]}
"""

    if provider == "OpenAI":
        raw = call_openai(api_key, model, prompt)
    elif provider == "Anthropic":
        raw = call_anthropic(api_key, model, prompt)
    elif provider == "Gemini":
        raw = call_gemini(api_key, model, prompt)
    else:
        raise ValueError("Unknown provider")

    try:
        data = json.loads(raw)
        slides = data.get("slides")
        if not isinstance(slides, list) or not slides:
            raise ValueError("Missing or empty 'slides' array")
        # normalize fields
        norm_slides = []
        for s in slides:
            title = str(s.get("title", "Untitled")).strip()
            bullets = s.get("bullets", [])
            if not isinstance(bullets, list):
                bullets = [str(bullets)]
            bullets = [str(b).strip() for b in bullets if str(b).strip()]
            entry = {"title": title, "bullets": bullets}
            if want_notes and isinstance(s.get("notes"), str):
                entry["notes"] = s["notes"].strip()
            norm_slides.append(entry)
        return LLMResponse(slides=norm_slides)
    except json.JSONDecodeError:
        raise RuntimeError("Provider did not return valid JSON. Try again or adjust your model.")


# ------------------------------
# Non-LLM fallback slide mapping
# ------------------------------
MD_H1 = re.compile(r"^# +(.+)$", re.MULTILINE)
MD_H2 = re.compile(r"^## +(.+)$", re.MULTILINE)


def heuristic_outline(text: str, want_notes: bool) -> LLMResponse:
    text = text.strip()
    slides: List[Dict[str, Any]] = []

    # Use markdown headings if present
    headings = MD_H1.findall(text) or MD_H2.findall(text)
    if headings:
        sections = re.split(r"^##? +.+$", text, flags=re.MULTILINE)
        for i, section in enumerate(sections):
            if not section.strip():
                continue
            title = headings[i] if i < len(headings) else f"Section {i+1}"
            bullets = [line.strip("- *\t ") for line in section.splitlines() if line.strip().startswith(('-', '*'))]
            if not bullets:
                # chunk into sentences
                bullets = [s.strip() for s in re.split(r"(?<=[.!?])\s+", section) if len(s.strip()) > 0][:5]
            slides.append({"title": title.strip(), "bullets": bullets[:7], **({"notes": ""} if want_notes else {})})
    else:
        # Split by paragraphs into 6-10 slides
        paras = [p.strip() for p in text.split('\n\n') if p.strip()]
        target = min(max(len(paras), 4), 10)
        chunks = max(1, len(paras) // target)
        for i in range(0, len(paras), chunks):
            blob = " ".join(paras[i:i+chunks])
            title = blob.split(". ")[0][:70] if blob else f"Slide {len(slides)+1}"
            bullets = [s.strip() for s in re.split(r"(?<=[.!?])\s+", blob) if s.strip()][:5]
            slides.append({"title": title or f"Slide {len(slides)+1}", "bullets": bullets[:7], **({"notes": ""} if want_notes else {})})

    if not slides:
        slides = [{"title": "Overview", "bullets": [text[:120]]}]

    return LLMResponse(slides=slides)


# ------------------------------
# Template analysis & PPT building
# ------------------------------
@dataclass
class TemplateAssets:
    picture_blobs: List[bytes] = field(default_factory=list)
    title_layout_idx: int = 0
    title_and_content_layout_idx: int = 0


def analyze_template(prs: Presentation) -> TemplateAssets:
    # Collect picture shapes from existing slides to reuse
    pics: List[bytes] = []
    for s in prs.slides:
        for shp in s.shapes:
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    pics.append(shp.image.blob)
                except Exception:
                    pass

    # Find candidate layouts
    title_idx = 0
    tac_idx = 1 if len(prs.slide_layouts) > 1 else 0
    for i, layout in enumerate(prs.slide_layouts):
        names = ",".join([ph.name.lower() for ph in layout.placeholders])
        lname = (getattr(layout, 'name', '') or '').lower()
        if ("title" in names or "title" in lname) and ("content" not in names):
            title_idx = i
        if ("title" in names and ("content" in names or "body" in names)):
            tac_idx = i
    return TemplateAssets(picture_blobs=pics, title_layout_idx=title_idx, title_and_content_layout_idx=tac_idx)


def add_title_slide(prs: Presentation, assets: TemplateAssets, title: str, subtitle: str = ""):
    layout = prs.slide_layouts[assets.title_layout_idx]
    slide = prs.slides.add_slide(layout)
    if slide.shapes.title:
        slide.shapes.title.text = title
        slide.shapes.title.text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
    # try subtitle placeholder
    for ph in slide.placeholders:
        if ph.is_placeholder and 'subtitle' in (ph.name or '').lower():
            ph.text = subtitle
            break
    _maybe_add_reused_picture(slide, assets)


def _maybe_add_reused_picture(slide, assets: TemplateAssets):
    # heuristically add one template image to bottom-right corner, if present
    if not assets.picture_blobs:
        return
    blob = assets.picture_blobs[hash(slide) % len(assets.picture_blobs)]
    try:
        left = slide.part.slide_width - Inches(3.0)
        top = slide.part.slide_height - Inches(2.2)
        slide.shapes.add_picture(io.BytesIO(blob), left, top, width=Inches(2.8))
    except Exception:
        pass


def add_content_slide(prs: Presentation, assets: TemplateAssets, title: str, bullets: List[str], notes: Optional[str] = None):
    layout = prs.slide_layouts[assets.title_and_content_layout_idx]
    slide = prs.slides.add_slide(layout)

    if slide.shapes.title:
        slide.shapes.title.text = title

    # find body placeholder
    body = None
    for ph in slide.placeholders:
        pname = (ph.name or '').lower()
        if 'content' in pname or 'body' in pname or 'text' in pname:
            body = ph
            break
    if body is None:
        # fallback: add a textbox
        left, top, width, height = Inches(1.0), Inches(1.8), Inches(8.0), Inches(4.5)
        body = slide.shapes.add_textbox(left, top, width, height).text_frame
    else:
        body = body.text_frame

    body.clear()
    for i, b in enumerate(bullets):
        p = body.add_paragraph() if i > 0 else body.paragraphs[0]
        p.text = b
        p.level = 0
        # Try to respect template font size; if missing, set a sane default
        try:
            p.font.size = p.font.size or Pt(18)
        except Exception:
            pass

    if notes is not None:
        try:
            slide.notes_slide.notes_text_frame.text = notes
        except Exception:
            # Some layouts may not support notes until created; python-pptx usually creates it lazily
            try:
                _ = slide.notes_slide
                slide.notes_slide.notes_text_frame.text = notes
            except Exception:
                pass

    _maybe_add_reused_picture(slide, assets)


# ------------------------------
# Streamlit UI
# ------------------------------

def _print_requirements_and_exit():
    print(REQUIREMENTS_TXT)
    sys.exit(0)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--print-reqs":
        _print_requirements_and_exit()

    _write_repo_files_once()

    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Paste text → upload template → pick provider → download PPT. Keys never stored.")

    with st.sidebar:
        st.subheader("LLM Provider")
        provider = st.selectbox("Choose provider", ["None (Heuristic)", "OpenAI", "Anthropic", "Gemini"], index=0)
        model = st.text_input("Model (optional)", help="e.g., gpt-4o-mini / claude-3-5-sonnet-latest / gemini-1.5-flash")
        api_key = st.text_input("API Key", type="password") if provider != "None (Heuristic)" else ""
        want_notes = st.checkbox("Auto-generate speaker notes", value=True)
        preset = st.selectbox("Tone preset (optional)", ["", "Professional", "Investor pitch", "Visual-heavy", "Technical", "Educator"], index=0)

    col1, col2 = st.columns([2, 1])
    with col1:
        user_text = st.text_area("Paste your text or markdown", height=320, placeholder="# Title\n\nPaste lots of text…")
        guidance = st.text_input("One-line guidance (optional)", value=(preset or ""))

    with col2:
        uploaded = st.file_uploader("Upload a PowerPoint template/presentation (.pptx or .potx)", type=["pptx", "potx"], accept_multiple_files=False)
        st.caption("We will reuse its layouts, colors, and any images when generating slides.")
        gen_btn = st.button("Generate Presentation", type="primary", use_container_width=True)

    if gen_btn:
        if not user_text or not uploaded:
            st.error("Please provide both input text and a PowerPoint template.")
            return
        size_mb = (uploaded.size or 0) / (1024 * 1024)
        if size_mb > MAX_UPLOAD_MB:
            st.error(f"Template file too large ({size_mb:.1f} MB). Limit: {MAX_UPLOAD_MB} MB.")
            return

        with st.spinner("Analyzing template & mapping slides…"):
            try:
                prs = Presentation(uploaded)
            except Exception as e:
                st.error(f"Failed to read PowerPoint file: {e}")
                return
            assets = analyze_template(prs)

            # Build outline
            try:
                if provider == "None (Heuristic)":
                    outline = heuristic_outline(user_text, want_notes)
                else:
                    if not api_key:
                        st.error("Enter your API key for the selected provider, or choose the heuristic mode.")
                        return
                    outline = outline_with_llm(provider, api_key, model, user_text, guidance, want_notes)
            except Exception as e:
                st.error(f"Outline generation failed: {e}")
                return

        # Build deck
        with st.spinner("Composing slides in your template style…"):
            try:
                # Start a fresh Presentation based on the same template to keep masters/theme
                template_bytes = uploaded.getvalue()
                prs_out = Presentation(io.BytesIO(template_bytes))

                # Optionally add a title slide from first outline entry if it looks like a title
                if outline.slides:
                    first = outline.slides[0]
                    add_title_slide(prs_out, assets, first.get("title", ""), guidance or "")
                    # Remove first from content if it’s only a title
                    remaining = outline.slides[1:] if (len(first.get("bullets", [])) <= 1) else outline.slides
                else:
                    remaining = []

                for s in remaining:
                    add_content_slide(prs_out, assets, s.get("title", "Untitled"), s.get("bullets", []), s.get("notes"))

                bio = io.BytesIO()
                prs_out.save(bio)
                bio.seek(0)
            except Exception as e:
                st.error(f"Failed to compose presentation: {e}")
                return

        st.success("Presentation ready!")
        st.download_button(
            label="Download .pptx",
            data=bio,
            file_name="generated_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

        with st.expander("Show JSON outline (debug)"):
            st.json({"slides": outline.slides})

    st.markdown("---")
    st.markdown(
        "**Privacy note:** Your API key (if provided) is used only to call the selected provider from your browser session to this server. It is never logged or stored."
    )


if __name__ == "__main__":
    main()
