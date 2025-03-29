from docx import Document
import re
from pathlib import Path

def convert_docx_to_text(docx_file, headers):
    doc = Document(docx_file)
    formatted_text = []
    current_section = None
    section_counter = 0
    
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
            
        text = para.text.strip()
        has_bold = any(run.bold for run in para.runs)
        
        if text in headers:
            section_counter += 1
            current_section = str(section_counter)
            processed_text = ""
            for run in para.runs:
                if run.bold:
                    processed_text += f"*{run.text}*"
                else:
                    processed_text += run.text
            formatted_text.append(f"{current_section}. {processed_text}")
            continue
        
        subsection_match = re.match(r'^([a-z])\.\s+(.*)', text)
        if subsection_match:
            subsection_letter = subsection_match.group(1)
            processed_text = ""
            for run in para.runs:
                if run.bold:
                    processed_text += f"*{run.text}*"
                else:
                    processed_text += run.text
            content_part = processed_text[2:].lstrip()
            formatted_text.append(f"   {subsection_letter}. {content_part}")
            continue
        
        if current_section and len(formatted_text) > 0 and formatted_text[-1].startswith(f"{current_section}."):
            subsection_letter = "a"
            processed_text = ""
            for run in para.runs:
                if run.bold:
                    processed_text += f"*{run.text}*"
                else:
                    processed_text += run.text
            formatted_text.append(f"   {subsection_letter}. {processed_text}")
            continue
            
        if current_section and len(formatted_text) > 0:
            last_line = formatted_text[-1]
            if last_line.startswith("   "):
                last_letter_match = re.match(r'\s+([a-z])\.\s+', last_line)
                if last_letter_match:
                    last_letter = last_letter_match.group(1)
                    next_letter = chr(ord(last_letter) + 1)
                    processed_text = ""
                    for run in para.runs:
                        if run.bold:
                            processed_text += f"*{run.text}*"
                        else:
                            processed_text += run.text
                    formatted_text.append(f"   {next_letter}. {processed_text}")
                    continue
        
        processed_text = ""
        for run in para.runs:
            if run.bold:
                processed_text += f"*{run.text}*"
            else:
                processed_text += run.text
        formatted_text.append(f"      {processed_text}")
    
    return formatted_text

def extract_bold_words(formatted_text):
    extracted_words = []
    current_section = None
    current_subsection = None
    
    for line in formatted_text:
        if not line.strip():
            continue
        
        section_match = re.match(r'^(\d+)\.\s+', line)
        if section_match:
            current_section = section_match.group(1)
            continue
        
        subsection_match = re.match(r'^\s*([a-z])\.\s+', line)
        if subsection_match:
            current_subsection = subsection_match.group(1)
        
        if current_section and current_subsection:
            bold_matches = re.findall(r'\*(.*?)\*', line)
            for bold_word in bold_matches:
                if bold_word.strip():
                    section_id = f"{current_section}{current_subsection}"
                    entry = f"{section_id}: {bold_word.strip()}"
                    extracted_words.append(entry)
    
    return extracted_words

docx_file_path = "/mnt/data/script.docx"
headers_list = [
    "Really?",
    "How is AI Modeled Today?",
    "Me",
    "Car analogy",
    "So What",
    "Neurochemistry",
    "The Skills That Keep You Relevant",
    "The Future: AI as a Tool, Not a Threat",
    "Closing"
]

formatted_text = convert_docx_to_text(docx_file_path, headers_list)
bold_words = extract_bold_words(formatted_text)

formatted_text_path = "/mnt/data/zfinal.txt"
bold_words_path = "/mnt/data/zbold.txt"

with open(formatted_text_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(formatted_text))

with open(bold_words_path, 'w', encoding='utf-8') as f:
    for word in bold_words:
        f.write(word + '\n')

formatted_text_path, bold_words_path