"""
media
"""

import os

def createMediaRel(slide, filename):
    mediaPath = os.path.join(os.path.dirname(__file__), "media")

    return slide.shapes.part.get_or_add_image_part(os.path.join(mediaPath, filename))

