from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
import os

prs = Presentation()
prs.slide_height = Inches(9)
prs.slide_width = Inches(16)

with open("./settings.txt", 'r', encoding='utf-8') as settings_file:
    settings = settings_file.readlines()
    for setting in settings:
        key, value = setting.strip().split('=')
        if key == "font_size":
            font_size = int(value)

for text in os.listdir("./text"):
    with open(os.path.join("./text", text), 'r', encoding='utf-8') as file:
        content = file.read()
        sloky = content.split("\n\n")

        for s in sloky:
            title_slide_layout = prs.slide_layouts[6] # Blank slide layout
            slide = prs.slides.add_slide(title_slide_layout)

            txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(16 - 2), Inches(9 - 2))
            tf = txBox.text_frame
            #tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

            p = tf.add_paragraph()
            p.font.size = Pt(font_size)

            p.text = s.strip()
            p.alignment = PP_ALIGN.CENTER
prs.save('presentation.pptx')
