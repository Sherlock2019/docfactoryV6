# 🧩 Rackspace Documentation Factory

This app lets Rackspace consultants create CRA, Solution, and SOW documents automatically using branded templates with placeholders like `{CUSTOMER_NAME}`, `{SLAS}`, etc.

## Features

- Upload or use predefined `.dotx` templates
- Auto-fill standard sections for CRA, Solution, and SOW
- Upload content or images per placeholder
- Supports `.docx`, `.pptx`, `.xlsx`, `.txt`, `.png`, `.jpg`
- Generate `.docx` or `.pptx` output with placeholder replacements
- PDF export ready
- Add and update new templates

## Usage

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Project Structure

```
docfactory_rackspace/
├── streamlit_app.py
├── templates/
│   ├── rackspace_template_CRA.dotx
│   ├── rackspace_template_Solution.dotx
│   └── rackspace_template_SOW.dotx
├── examples/
├── assets/screenshots/
├── requirements.txt
└── README.md
```
