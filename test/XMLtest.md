template: Martin Template.pptx
bookmark: 2
aapplescript: test.scpt
aapplescriptOptions: run reload

### Bulleted List

* One
* Two
* Three

``` run-python

from media import *
image_part0, rId0 = createMediaRel(slide, "cross-black.png")
image_part1, rId1 = createMediaRel(slide, "tick-black.png")
image_part2, rId2 = createMediaRel(slide, "partial-black.png")

tf = slide.shapes[-1].text_frame
paras = tf.paragraphs
for paraNumber, para in enumerate(paras):
    if paraNumber == 0:
      rId = rId0
    elif paraNumber == 1:
      rId = rId1
    else:
      rId = rId1

    originalFontSize = para.font.size
    
    # Save original indentation level
    level = para.level
    
    # Remove the original pPr element
    para._element.remove(para._element.getchildren()[0])
    
    xml = ''
    
    # Note the level insertion
    xml += f'<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" lvl="{level}">'
    
    
    xml += '<a:buBlip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    xml += f'    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rId}"/>'
    xml += '</a:buBlip>'

    xml += '</a:pPr>'
    
    RunPython.attachXMLfile(para._element, "/Users/martinpacker/md2pptx/test/XMLtest")
    
    # Restore original font size
    para.font.size = originalFontSize

```