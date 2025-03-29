from docx import Document
import sys
import re
from pathlib import Path

def process_document(docx_file, formatted_output_file, bold_words_output_file, headers):
    try:
        print(f"Processing document: {docx_file}")
        
        formatted_text = convert_docx_to_text(docx_file, headers)
        
        with open(formatted_output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(formatted_text))
        
        print(f"Successfully converted document to formatted text")
        print(f"Formatted text saved to {formatted_output_file}")
        
        extracted_words = extract_bold_words(formatted_text)
        
        with open(bold_words_output_file, 'w', encoding='utf-8') as f:
            for entry in extracted_words:
                f.write(entry + '\n')
        
        print(f"Successfully extracted {len(extracted_words)} bold words")
        print(f"Bold words saved to {bold_words_output_file}")
        
        if extracted_words:
            section_counts = {}
            for entry in extracted_words:
                section_id = entry.split(':')[0].strip()
                section = section_id[0]
                if section not in section_counts:
                    section_counts[section] = 0
                section_counts[section] += 1
            
            print("\nDistribution by section:")
            for section, count in sorted(section_counts.items()):
                print(f"Section {section}: {count} words")
        
        return True
        
    except Exception as e:
        print(f"Error processing document: {e}")
        import traceback
        traceback.print_exc()
        return False

def convert_docx_to_text(docx_file, headers):
    doc = Document(docx_file)
    
    formatted_text = []
    
    current_section = None
    
    section_counter = 0
    
    for i, para in enumerate(doc.paragraphs):
        if not para.text.strip():
            continue
            
        text = para.text.strip()
        
        has_bold = any(run.bold for run in para.runs)
        
        if any(keyword in text for keyword in headers):
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
            subsection_content = subsection_match.group(2)
            
            processed_text = ""
            for run in para.runs:
                if run.bold:
                    processed_text += f"*{run.text}*"
                else:
                    processed_text += run.text
            
            letter_part = processed_text[:2]
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

def main():
    if len(sys.argv) < 3:
        print("Usage: python working2.py input_file.docx header1 [header2 ...]")
        print("Output files will be zfinal.txt and zbold.txt")
        print("At least one header is required")
        sys.exit(1)
    
    docx_file = sys.argv[1]
    formatted_output_file = "zfinal.txt"
    bold_words_output_file = "zbold.txt"
    
    if not Path(docx_file).exists():
        print(f"Error: Input file '{docx_file}' does not exist")
        sys.exit(1)
    
    # Get headers (now required)
    headers = sys.argv[2:]
    print(f"Using headers: {', '.join(headers)}")
    
    if not process_document(docx_file, formatted_output_file, bold_words_output_file, headers):
        sys.exit(1)

if __name__ == "__main__":
    main() 