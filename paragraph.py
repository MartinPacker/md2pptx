"""
paragraph
"""


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