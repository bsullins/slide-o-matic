from pptx import Presentation

prs = Presentation('pluralsight-template.pptx')


layout_title = prs.slide_layouts[0]
layout_subheader = prs.slide_layouts[2]

title_slide = prs.slides.add_slide(layout_title)
title = title_slide.shapes.title

section_1 = prs.slides.add_slide(layout_subheader)
section_title= section_1.shapes.title

title.text = "Making Presentations from Python!"
section_title.text = "This is where it get's good!"

prs.save('ps-test-2.pptx')