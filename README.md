# Presentation-Preparer

# Your Text, Your Style – Auto-Generate a Presentation


A tiny web app that turns bulk text/markdown into a PowerPoint that matches an uploaded template’s look & feel. Bring your own LLM API key (OpenAI/Anthropic/Gemini). No keys are stored.


## Features
- Paste long text or markdown
- Optional one-line guidance (e.g., "investor pitch deck")
- Upload a .pptx/.potx template to copy styles, layouts, fonts, and reuse images
- Choose LLM provider and model; provide your API key (not stored)
- Download a new .pptx (no AI images created)


## Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```


## Requirements
See `requirements.txt` or run `python app.py --print-reqs`.


## How It Works (200–300 words)
The app first gathers user input: a block of text (or markdown), optional guidance, the preferred LLM provider/model with an API key, and a PowerPoint template. If the user does not provide an API key, the app falls back to a deterministic heuristic: it creates a title slide and content slides by splitting the input into sections using markdown headings or length-based chunking.


When an LLM key is provided, the app prompts the model to return a strict JSON outline containing slide objects with `title`, `bullets`, and optional `notes`. This outline determines the number and structure of slides. The JSON is validated before use. The app never logs or stores the API key; it is kept in-memory for the single request.


To apply visual style, the app instantiates `Presentation(template_file)` via `python-pptx`, so all new slides are added by reusing the template’s masters and layouts. It attempts to pick a title/content layout for most slides and copies paragraph-level properties (font size, bold, bullet levels) when possible. The app scans the uploaded template for picture shapes and reuses a few of them as decorative images on generated slides to preserve the brand feel, without creating any AI images.


Finally, the app writes slides using the chosen layout, sets title and bullets, optionally adds speaker notes from the outline, and offers a download button for the resulting `.pptx`.
