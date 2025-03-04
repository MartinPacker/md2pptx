"""
card
"""

myVersion = "0.1"

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
        self.graphicTitle = None

        self.printableFilename = None
        
        # Both audio and video
        self.mediaInfo = None
        self.mediaDimensions = None
        self.mediaShape = None
                
        self.mediaURL = None
        
        self.backgroundShape = None
        self.backgroundTop = None
        
        self.bodyShape = None
        self.bodyTop = None
        
        self.top = None
        self.left = None

