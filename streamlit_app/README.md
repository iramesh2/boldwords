# Document Bold Text Extractor

A Streamlit application that extracts bold text from Word documents, organized by automatically identified sections.

## Features

- Upload Word documents (.docx)
- Automatically identify section headers using OpenAI's GPT-4o model
- Extract and organize bold text by sections and subsections
- Display formatted document structure
- Show extracted bold text in organized tabs
- Download results as JSON
- Provide a table view of all extracted terms

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Set up your OpenAI API key as an environment variable named `API`:
   ```
   export API=your_openai_api_key
   ```

## Running the Application

Run the Streamlit app with:

```
streamlit run app.py
```

## Usage

1. Upload a Word document using the file uploader
2. The app will:
   - Convert the document to text
   - Use OpenAI to identify section headers
   - Process the document to extract bold text by sections
   - Display the results in various formats

## How It Works

1. The document is first processed to extract raw text
2. The text is sent to OpenAI's GPT-4o model to identify section headers
3. The document is processed again using the identified headers
4. Bold text is extracted and organized by section/subsection
5. The results are displayed in the Streamlit interface

## Requirements

- Python 3.7+
- OpenAI API key
- Required Python packages (see requirements.txt) 