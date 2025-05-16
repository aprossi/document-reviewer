# DOCX Feedback Tools with Ollama LLMs

This repository contains Python scripts for providing stylistic feedback on Word (.docx) and OpenDocument Text (.odt) files using Large Language Models via Ollama.

## Scripts

- **docx_comments.py** - Adds side comments to Word documents with feedback
- **odt_comments.py** - Adds side comments to OpenDocument Text (.odt) files with feedback

The repository also includes an example prompt file (`proofreading-english-instructions.txt`) for English language proofreading.

## Features

- Overall document style assessment
- Detailed text element analysis
- Support for tables and text boxes
- Separate summary document (optional)
- Works with any LLM model available in Ollama
- Adjustable creativity level
- Custom system prompts

## Requirements

- Python 3.8+
- Ollama installed and running
- Word document (.docx) or OpenDocument (.odt) files

## Installation

### 1. Install Ollama

#### macOS (with Homebrew)
```bash
brew install ollama
```

#### Other platforms
Download from [ollama.com/download](https://ollama.com/download)

### 2. Start Ollama
```bash
ollama serve
```

### 3. Pull a model
```bash
ollama pull llama3.1  # Recommended starting model
# or
ollama pull mistral   # Alternative model
```

### 4. Set up Python environment

```bash
# Clone the repository
git clone https://github.com/aprossi/document-reviewer.git
cd document-reviewer

# Create virtual environment
python3 -m venv venv

# Activate on Windows
venv\Scripts\activate

# Activate on macOS/Linux
source venv/bin/activate

# upgrade pip
pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Word Document Feedback

```bash
python docx_comments.py your_document.docx
```

### OpenDocument Text Feedback

```bash
python odt_comments.py your_document.odt
```

### Common Options

```bash
# List available models
python docx_comments.py --list-models

# Use a specific model
python docx_comments.py your_document.docx --model mistral

# More creative feedback
python docx_comments.py your_document.docx --creative

# Custom system prompt
python docx_comments.py your_document.docx --system-prompt my_prompt.txt

# Verbose progress
python docx_comments.py your_document.docx --verbose

# Custom Ollama API host
python docx_comments.py your_document.docx --api-host http://localhost:11434

# Disable summary document
python docx_comments.py your_document.docx --no-summary
```

## Output Files

Each script produces:
1. A document with feedback comments in the margins
2. A summary document with all feedback in one place (optional)

## Customization

Create custom system prompts by saving them as text files. The repository includes an example for English proofreading:

### English Proofreading

To use the included English proofreading prompt:

```bash
python docx_comments.py your_document.docx --system-prompt english_proofreading_prompt.txt
```

The prompt focuses on identifying:
- Grammar and syntax errors
- Spelling mistakes and typos
- Punctuation usage
- Subject-verb agreement
- Tense consistency
- Article and preposition usage
- Word choice and clarity

You can create your own custom prompts by following a similar format:

## Troubleshooting

- Make sure Ollama is running with `ollama serve`
- Verify available models with `--list-models`
- For model-specific issues, try a different model
- If comments don't appear correctly, check the summary document

## License

MIT
