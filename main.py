import pypandoc
from pathlib import Path
import re

cwd = Path().resolve()
input_directory = cwd / 'documents'
output_directory = cwd / 'markdown'

def convert_files():
  for docx_file in Path(input_directory).glob("*.docx"):

    # Convert DOCX to Markdown
    md_file = Path(output_directory) / docx_file.with_suffix('.md').name
    pypandoc.convert_file(str(docx_file), 'md', outputfile=str(md_file))

    clean_file(md_file)

def clean_file(file_path):
  with open(file_path, 'r', encoding='utf-8') as file:
    content = file.read()
  
  # Removes { .unnumbered} headings from docx > md parse.
  pattern = r'\s*\{#.*?\.unnumbered\}'
  cleaned_content = re.sub(pattern, '', content)
  
  with open(file_path, 'w', encoding='utf-8') as file:
    file.write(cleaned_content)

if __name__ == '__main__':
  convert_files()


