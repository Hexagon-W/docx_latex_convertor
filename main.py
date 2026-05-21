import argparse
import sys
import re
from pathlib import Path

# Safeguard check to ensure docxlatex is installed in the active environment
try:
    from docxlatex import Document
except ImportError:
    print("❌ Error: Module 'docxlatex' not found.")
    print("👉 Install it with: pip install docxlatex")
    print(f"Running Python executable: {sys.executable}")
    sys.exit(1)


def extract_equations(docx_path: Path, output_path: Path, output_format: str) -> bool:
    """Extracts equations from a .docx file and writes them to a text file.
    
    Supports exporting the full text structure or isolating individual equations.
    """
    if not docx_path.exists():
        print(f"❌ Error: DOCX file not found at '{docx_path}'")
        return False

    try:
        # Initialize docxlatex Document (expects a string path string)
        doc = Document(str(docx_path))
        full_text = doc.get_text()
    except Exception as e:
        print(f"❌ Error opening or parsing document: {e}")
        return False

    try:
        with output_path.open('w', encoding='utf-8') as f:
            if output_format == 'full':
                # Format 1: Write the entire document layout with inline ($) and block ($$) markers
                f.write(full_text)
                print(f"✅ Successfully wrote full document text to: {output_path}")
            
            elif output_format == 'split':
                # Format 2: Isolate and list out just the mathematical equations using regex
                block_pattern = re.compile(r"\$\$(.*?)\$\$", re.DOTALL)
                inline_pattern = re.compile(r"(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)", re.DOTALL)
                
                # Find display blocks, then strip them to avoid matching inner contents as inlines
                blocks = block_pattern.findall(full_text)
                text_no_blocks = block_pattern.sub("", full_text)
                inlines = inline_pattern.findall(text_no_blocks)
                
                equations = []
                for b in blocks:
                    equations.append({'type': 'display/block', 'content': b.strip()})
                for i in inlines:
                    equations.append({'type': 'inline', 'content': i.strip()})
                
                # Write a clean, structured manifest summary
                f.write(f"Source Document: {docx_path.name}\n")
                f.write(f"Total Equations Found: {len(equations)}\n")
                f.write("=" * 40 + "\n\n")
                
                for idx, eq in enumerate(equations, start=1):
                    f.write(f"--- EQUATION {idx} ({eq['type']}) ---\n")
                    f.write(f"{eq['content']}\n\n")
                
                print(f"✅ Extracted {len(equations)} isolated equations to: {output_path}")
                
        return True
    except IOError as e:
        print(f"❌ Failed to write output file: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description='Extract LaTeX math equations from a Microsoft Word (.docx) document.'
    )
    
    # Using nargs='?' makes the document argument optional.
    # If omitted, it defaults to looking for 'Newmark_Newton_Raphson_Algorithm.docx' in your current folder.
    parser.add_argument(
        'docx', 
        type=str, 
        nargs='?', 
        default='Newmark_Newton_Raphson_Algorithm.docx',
        help="Path to the source .docx file (default: 'Newmark_Newton_Raphson_Algorithm.docx')"
    )
    
    parser.add_argument(
        '-o', '--output', 
        type=str, 
        default='extracted_math.txt',
        help="Path to the output text file (default: 'extracted_math.txt')"
    )
    
    parser.add_argument(
        '--format', 
        choices=['full', 'split'], 
        default='full',
        help="Output format: 'full' saves layout text with math syntax. "
             "'split' isolates and lists numbered equations sequentially."
    )
    
    args = parser.parse_args()
    
    docx_path = Path(args.docx)
    output_path = Path(args.output)
    
    success = extract_equations(docx_path, output_path, args.format)
    if not success:
        sys.exit(1)


if __name__ == '__main__':
    main()