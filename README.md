Markdown to Word Converter
Professional Python script for automated conversion of Markdown (.md) files into formatted Word (.docx) documents while preserving document structure and applying consistent styling from Word templates.

Features
Template-Based Styling: Apply organizational styles and formatting from existing Word templates

Comprehensive Markdown Support:

Headers (H1-H6)

Bold, italic, and mixed inline formatting

Code blocks with syntax highlighting preservation

Inline code

Ordered and unordered lists with nesting

Blockquotes

Horizontal rules

Links

Paragraphs and line breaks

Intelligent Style Mapping: Automatically maps markdown elements to Word styles

Fallback Handling: Gracefully handles missing styles with sensible defaults

Robust Error Handling: Clear error messages for missing files, encoding issues, or permission problems

Flexible Output: Specify custom output locations or use automatic naming

Professional Output: Console progress indicators and detailed status reporting

Installation
Prerequisites
Python 3.7 or higher

pip (Python package manager)

Install Dependencies
bash
pip install -r requirements.txt
Or install packages individually:

bash
pip install python-docx>=0.8.11
pip install mistune>=2.0.5
Usage
Basic Conversion
Convert a markdown file using default formatting:

bash
python md_to_word.py document.md
Output: document_converted.docx (in the same directory as input)

With Word Template
Apply styles from an existing Word template:

bash
python md_to_word.py document.md -t company_template.docx
Specify Output Location
Define custom output path and filename:

bash
python md_to_word.py document.md -o reports/final_report.docx
Full Command with All Options
bash
python md_to_word.py document.md -t template.docx -o output/report.docx -v
Command-Line Arguments
| Argument

# convertMDtoWord
