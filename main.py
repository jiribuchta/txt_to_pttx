from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
import os

prs = Presentation()
prs.slide_height = Inches(9)
prs.slide_width = Inches(16)

if not os.path.exists("./text"):
    os.makedirs("./text")

if not os.path.exists("./settings.txt"):
    with open("./settings.txt", 'w', encoding='utf-8') as settings_file:
        settings_file.write("font_size=72\n")

if os.listdir("./text") == []:
    print("Please add text files to the 'text' directory and run the script again.")
    input("Press Enter to exit...")

with open("./settings.txt", 'r', encoding='utf-8') as settings_file:
    settings = settings_file.readlines()
    for setting in settings:
        key, value = setting.strip().split('=')
        if key == "font_size":
            font_size = int(value)

for text in os.listdir("./text"):
    with open(os.path.join("./text", text), 'r', encoding='utf-8') as file:
        content = file.read()
        verse = content.split("\n\n")

        for s in verse:
            title_slide_layout = prs.slide_layouts[6] # Blank slide layout
            slide = prs.slides.add_slide(title_slide_layout)

            txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(16 - 2), Inches(9 - 2))
            tf = txBox.text_frame

            p = tf.add_paragraph()
            p.font.size = Pt(font_size)

            p.text = s.strip()
            p.alignment = PP_ALIGN.CENTER
prs.save('presentation.pptx')
