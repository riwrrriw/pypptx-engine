"""
pypptx-engine: JSON to PowerPoint presentation generator
"""
from .engine import PPTXEngine
from .slides import SlideManager
from .shapes import ShapeFactory
from .formatters import FontFormatter, LineFormatter, ShadowFormatter, ColorFormatter
from .flowchart import FlowchartHandler, FlowchartLayoutManager
from .templates import TemplateManager
from .animations import AnimationManager, TransitionPresets, AnimationPresets

__version__ = "0.1.0"
__all__ = [
    'PPTXEngine',
    'SlideManager', 
    'ShapeFactory',
    'FontFormatter',
    'LineFormatter',
    'ShadowFormatter',
    'ColorFormatter',
    'FlowchartHandler',
    'FlowchartLayoutManager',
    'TemplateManager',
    'AnimationManager',
    'TransitionPresets',
    'AnimationPresets'
]