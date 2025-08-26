"""
pypptx-engine: JSON to PowerPoint presentation generator
"""
from .engine import PPTXEngine
from .slides import SlideManager
from .shapes import ShapeFactory
from .formatters import ColorFormatter, FontFormatter, LineFormatter, ShadowFormatter

__version__ = "0.1.0"
__all__ = [
    "PPTXEngine",
    "SlideManager", 
    "ShapeFactory",
    "ColorFormatter",
    "FontFormatter", 
    "LineFormatter",
    "ShadowFormatter"
]