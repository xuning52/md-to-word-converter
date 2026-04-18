#pip install pypandoc python-docx
import pypandoc
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def set_format(paragraph, font_east, font_west, size, is_bold=False, align=None, is_title=False):
    """
    Applies precise formatting to each paragraph.
    Titles: 1.5x spacing | Body: 1.25x spacing
    """
    if align:
        paragraph.alignment = align
    
    # Line spacing logic: 1.5 for titles, 1.25 for body text
    paragraph.paragraph_format.line_spacing = 1.5 if is_title else 1.25

    if not paragraph.runs:
        paragraph.add_run()
    
    for run in paragraph.runs:
        run.font.size = Pt(size)
        run.font.bold = is_bold
        # Force Black color for all text
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Set Western Font (Times New Roman)
        run.font.name = font_west
        # Set East Asian Font (SimSun or Heiti)
        r = run._element.get_or_add_rPr()
        r_fonts = r.get_or_add_rFonts()
        r_fonts.set(qn('w:eastAsia'), font_east)

def apply_custom_styles(docx_path):
    """
    Post-processing the Word document to inject specific fonts and spacing.
    """
    doc = Document(docx_path)
    
    for para in doc.paragraphs:
        style_name = para.style.name
        
        if style_name == 'Heading 1':
            # H1: 22pt (Level 2), Heiti, Center, Bold, 1.5x
            set_format(para, '黑体', 'Times New Roman', 22, True, WD_ALIGN_PARAGRAPH.CENTER, is_title=True)
        elif style_name == 'Heading 2':
            # H2: 15pt (Small 3), SimSun, Left, Bold, 1.5x
            set_format(para, '宋体', 'Times New Roman', 15, True, WD_ALIGN_PARAGRAPH.LEFT, is_title=True)
        elif style_name == 'Heading 3':
            # H3: 12pt (Small 4), SimSun, Left, Bold, 1.5x
            set_format(para, '宋体', 'Times New Roman', 12, True, WD_ALIGN_PARAGRAPH.LEFT, is_title=True)
        else:
            # Body: 10.5pt (Level 5), SimSun, Left, 1.25x
            set_format(para, '宋体', 'Times New Roman', 10.5, False, WD_ALIGN_PARAGRAPH.LEFT, is_title=False)
            
    doc.save(docx_path)

def convert_file(input_path):
    """
    Core conversion logic using Pandoc + Docx refinement.
    """
    output_path = input_path.rsplit('.', 1)[0] + '.docx'
    # Keep math support for Quantum/Physics notes
    extra_args = ['--mathjax', '--from=markdown+tex_math_dollars']
    
    try:
        pypandoc.convert_file(input_path, 'docx', outputfile=output_path, extra_args=extra_args)
        apply_custom_styles(output_path)
        print(f"  [SUCCESS] Created & Formatted: {os.path.basename(output_path)}")
    except Exception as e:
        print(f"  [FAILED] {os.path.basename(input_path)} | Reason: {e}")

def main():
    try:
        version = pypandoc.get_pandoc_version()
        print(f"--- [SYSTEM] Pandoc Environment: OK (Version {version}) ---")
    except OSError:
        print("\n[!] ERROR: Pandoc not found in System Path.")
        return

    print("\n" + "="*50)
    print("      MARKDOWN TO WORD CONVERTER (v3.2 PRO)")
    print("      Styles: Songti/TNR | H:1.5x | Body:1.25x")
    print("="*50)
    print("Choose your conversion mode:")
    print("  1. SINGLE FILE (Convert 1 specific .md file)")
    print("  2. ALL FILES   (Convert every .md file in a folder)")
    print("-" * 50)
    
    choice = input("Enter selection (1 or 2): ").strip()

    if choice == '1':
        print("\n[MODE: Single File]")
        path = input(">>> Please drag/paste the [FILE] path here: ").strip('"').strip()
        if os.path.isfile(path) and path.lower().endswith('.md'):
            convert_file(path)
        else:
            print("Invalid input! Make sure it is a valid .md file path.")

    elif choice == '2':
        print("\n[MODE: All Files in Folder]")
        folder = input(">>> Please drag/paste the [FOLDER] path here: ").strip('"').strip()
        if os.path.isdir(folder):
            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.md')]
            if not files:
                print("No Markdown (.md) files found in this directory.")
            else:
                print(f"Processing {len(files)} files...")
                for f in files:
                    convert_file(f)
        else:
            print("Invalid directory path!")
    else:
        print("Selection error! Please run the script again and enter 1 or 2.")

    print("\n" + "="*50)
    print("Task completed. Have a great day!")
    input("Press [Enter] to close this window...")

if __name__ == "__main__":
    main()