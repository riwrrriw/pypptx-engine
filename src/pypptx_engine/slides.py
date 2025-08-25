"""
Slide management and layout handling
"""
from __future__ import annotations

from typing import Any, Dict

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


class SlideManager:
    """Manage slide creation and layout operations."""
    
    def __init__(self, color_formatter):
        self.color_formatter = color_formatter
    
    def create_slide(self, prs: Presentation, slide_config: Dict[str, Any], base_dir: str, shape_factory) -> None:
        """Create a slide with specified layout and content."""
        layout_index = slide_config.get("layout", 6)  # Default to blank layout
        slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
        
        # Apply slide background
        self._apply_background(slide, slide_config.get("background"))
        
        # Add shapes to slide
        shapes_config = slide_config.get("shapes", [])
        for shape_config in shapes_config:
            shape_factory.create_shape(slide, shape_config, base_dir)
        
        # Handle placeholders if specified
        placeholders_config = slide_config.get("placeholders", {})
        self._fill_placeholders(slide, placeholders_config)
    
    def _apply_background(self, slide, background_config) -> None:
        """Apply background formatting to slide."""
        if not background_config:
            return
        
        if isinstance(background_config, str):
            # Simple color background
            color = self.color_formatter.parse_color(background_config)
            if color:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = color
        elif isinstance(background_config, dict):
            fill = slide.background.fill
            bg_type = background_config.get("type", "solid")
            
            if bg_type == "solid":
                fill.solid()
                color = self.color_formatter.parse_color(background_config.get("color"))
                if color:
                    fill.fore_color.rgb = color
            elif bg_type == "gradient":
                fill.gradient()
                # Gradient implementation would go here
            elif bg_type == "picture":
                # Picture background implementation
                pass
    
    def _fill_placeholders(self, slide, placeholders_config: Dict[str, Any]) -> None:
        """Fill slide placeholders with content."""
        if not placeholders_config:
            return
        
        for placeholder in slide.placeholders:
            placeholder_name = placeholders_config.get(str(placeholder.placeholder_format.idx))
            if placeholder_name and placeholder_name in placeholders_config:
                content = placeholders_config[placeholder_name]
                
                if placeholder.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                    if hasattr(placeholder, 'text_frame'):
                        placeholder.text = content.get("text", "")
                    elif hasattr(placeholder, 'insert_picture'):
                        # Handle picture placeholders
                        image_path = content.get("image_path")
                        if image_path:
                            placeholder.insert_picture(image_path)
