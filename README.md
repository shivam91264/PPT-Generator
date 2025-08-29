# Auto-PPT Generator

Auto-PPT Generator is a web application that transforms bulk user text (markdown or prose) into fully styled PowerPoint presentations. Using a user-uploaded PowerPoint template, it applies the template’s visual style, including layouts, fonts, colors, and images, to produce polished presentations automatically.

## Features

- Paste or upload large chunks of text for slide generation
- Upload any PowerPoint template (.pptx or .potx) for styling
- Choose to use LLM-powered slide structuring with your own API key (optional)
- Heuristic fallback mode for quick, no-key slide generation
- Reuses images from the uploaded template (no AI-generated images)
- Download the generated .pptx file matching your brand/style

## Setup and Usage

### Prerequisites

- Python 3.8+
- pip package manager

### Installation

1. Clone the repository:


2. Install dependencies:
pip install -r requirements.txt



### Running the App

Start the Flask development server:

python app.py