"""
card
"""

myVersion = "0.0"

__version__ = myVersion

class Card:
    def __init__(
        self,
    ):
        self.title = ""
        self.titleShape = None
        
        self.bullets = ""
        
        self.graphic = ""
        self.graphicShape = None
        self.graphicDimensions = None
        
        self.backgroundShape = None
        self.backgroundTop = None
        
        self.bodyShape = None
        self.bodyTop = None
        
        self.top = None
        self.left = None
