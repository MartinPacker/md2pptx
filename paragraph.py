"""
paragraph
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
from lxml import etree

def removeBullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.insert(
        0,
        etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"),
    )


def removeBullets(textFrame):
    for p in textFrame.paragraphs:
        removeBullet(p)


def removeSelectedBullets(textFrame, removalArray):
    for bulletNumber in removalArray:
        removeBullet(textFrame.paragraphs[bulletNumber])

def findTitleShape(slide):
    if slide.shapes.title == None:
        # Have to use first shape as title
        return slide.shapes[0]
        
    else:
        return slide.shapes.title

def getParagraphs(slide, wantedParagraphs = []):
    paragraphTree = []
    for theShape in slide.shapes:
        if (theShape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX) | (theShape.placeholder_format.type == PP_PLACEHOLDER.OBJECT):
            paragraphTree.append(theShape.text_frame.paragraphs)
    return[ paragraphTree[0][i] for i in wantedParagraphs]