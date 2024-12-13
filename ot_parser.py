import re
import argparse
from docx import Document

def is_centered(paragraph):
    return paragraph.alignment == 1  # 1 indicates centered alignment

def process_centered_line(line):
    return f"<title> {line} </title>"

def process_chapter_line(line):
    match = re.match(r'([\u0700-\u074F]+)\[', line)
    if match:
        return f'<chapter no="{match.group(1)}">'
    return line

def process_heading_paragraph(paragraph):
    return f"<heading> {paragraph} </heading>"

def process_verse_paragraph(paragraph):
    sentences = re.split(r'(\d+)', paragraph)
    result = ""
    for i in range(1, len(sentences), 2):
        number = sentences[i]
        sentence = sentences[i + 1].strip()
        result += f'<verse no="{number}"> {sentence} </verse>'
    return result

def process_paragraph(paragraph):
    text = paragraph.text.strip()
    if is_centered(paragraph):
        return process_centered_line(text)
    elif re.match(r'[\u0700-\u074F]+\[', text):
        return process_chapter_line(text)
    elif not any(char.isdigit() for char in text):
        return process_heading_paragraph(text)
    elif text[0].isdigit():
        return process_verse_paragraph(text)
    return text

def process_document(doc_path, output_path):
    doc = Document(doc_path)
    output_lines = ["<doc>"]
    inside_chapter = False
    last_heading = None

    for paragraph in doc.paragraphs:
        processed_text = process_paragraph(paragraph)
        if processed_text.startswith('<chapter no='):
            if inside_chapter:
                output_lines.append("</chapter>")
            inside_chapter = True
            if last_heading:
                output_lines.append(last_heading)
                last_heading = None
            output_lines.append(processed_text)
        elif processed_text.startswith('<heading>'):
            if inside_chapter:
                last_heading = processed_text
            else:
                output_lines.append(processed_text)
        else:
            if last_heading:
                output_lines.append(last_heading)
                last_heading = None
            output_lines.append(processed_text)

    if inside_chapter:
        output_lines.append("</chapter>")
    if last_heading:
        output_lines.append(last_heading)
    output_lines.append("</doc>")
    output_path = output_path.replace('.docx', '.xml')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process a Word document.')
    parser.add_argument('-f', '--file', required=True, help='Input Word document file name')
    parser.add_argument('-o', '--output', required=True, help='Output file path')
    args = parser.parse_args()

    doc_path = args.file
    output_path = args.output
    process_document(doc_path, output_path)