"""
Formatting utilities for colors, fonts, fills, and other visual properties
"""
from __future__ import annotations

from typing import Any, Dict, Optional, Union

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR, MSO_FILL_TYPE
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT
from pptx.util import Pt, Inches


class ColorFormatter:
    """Handle color conversions and formatting."""
    
    @staticmethod
    def parse_color(color_spec: Union[str, Dict[str, Any], None]) -> Optional[RGBColor]:
        """Parse various color specifications into RGBColor."""
        if not color_spec:
            return None
        
        if isinstance(color_spec, str):
            return ColorFormatter._hex_to_rgb(color_spec)
        elif isinstance(color_spec, dict):
            if "hex" in color_spec:
                return ColorFormatter._hex_to_rgb(color_spec["hex"])
            elif "rgb" in color_spec:
                rgb = color_spec["rgb"]
                return RGBColor(rgb[0], rgb[1], rgb[2])
        
        return None
    
    @staticmethod
    def _hex_to_rgb(hex_color: str) -> RGBColor:
        """Convert hex color string to RGBColor."""
        if not hex_color:
            return RGBColor(0, 0, 0)
        
        hex_color = hex_color.strip().lstrip("#")
        if len(hex_color) != 6:
            return RGBColor(0, 0, 0)
        
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return RGBColor(r, g, b)
        except ValueError:
            return RGBColor(0, 0, 0)
    
    @staticmethod
    def apply_fill(shape, fill_config: Dict[str, Any]) -> None:
        """Apply fill formatting to a shape."""
        if not fill_config:
            return
        
        fill = shape.fill
        fill_type = fill_config.get("type", "solid")
        
        if fill_type == "solid":
            fill.solid()
            color = ColorFormatter.parse_color(fill_config.get("color"))
            if color:
                fill.fore_color.rgb = color
        elif fill_type == "gradient":
            # Basic gradient support
            fill.gradient()
            gradient_stops = fill_config.get("stops", [])
            for stop in gradient_stops:
                color = ColorFormatter.parse_color(stop.get("color"))
                if color:
                    # Note: python-pptx gradient API is limited
                    pass
        elif fill_type == "pattern":
            fill.patterned()
            # Pattern support would need more implementation
        elif fill_type == "picture":
            # Picture fill would need image path
            pass
        elif fill_type == "none":
            fill.background()


class FontFormatter:
    """Handle font formatting and text properties."""
    
    @staticmethod
    def apply_font_formatting(font, font_config: Dict[str, Any]) -> None:
        """Apply font formatting from configuration."""
        if not font_config:
            return
        
        if "name" in font_config:
            font.name = font_config["name"]
        
        if "size" in font_config:
            font.size = Pt(font_config["size"])
        
        if "bold" in font_config:
            font.bold = bool(font_config["bold"])
        
        if "italic" in font_config:
            font.italic = bool(font_config["italic"])
        
        if "underline" in font_config:
            font.underline = bool(font_config["underline"])
        
        if "color" in font_config:
            color = ColorFormatter.parse_color(font_config["color"])
            if color:
                font.color.rgb = color
    
    @staticmethod
    def apply_paragraph_formatting(paragraph, para_config: Dict[str, Any]) -> None:
        """Apply paragraph-level formatting."""
        if not para_config:
            return
        
        alignment = para_config.get("alignment", "").lower()
        if alignment == "left":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        elif alignment == "center":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        elif alignment == "right":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.RIGHT
        elif alignment == "justify":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        if "space_before" in para_config:
            paragraph.space_before = Pt(para_config["space_before"])
        
        if "space_after" in para_config:
            paragraph.space_after = Pt(para_config["space_after"])
        
        if "line_spacing" in para_config:
            paragraph.line_spacing = para_config["line_spacing"]


class LineFormatter:
    """Handle line and border formatting."""
    
    @staticmethod
    def apply_line_formatting(line, line_config: Dict[str, Any]) -> None:
        """Apply line formatting from configuration."""
        if not line_config:
            return
        
        if "color" in line_config:
            color = ColorFormatter.parse_color(line_config["color"])
            if color:
                line.color.rgb = color
        
        if "width" in line_config:
            line.width = Pt(line_config["width"])
        
        # Dash style, arrow heads, etc. could be added here


class ShadowFormatter:
    """Handle shadow effects."""
    
    @staticmethod
    def apply_shadow(shape, shadow_config: Dict[str, Any]) -> None:
        """Apply shadow formatting to a shape."""
        if not shadow_config:
            return
        
        shadow = shape.shadow
        shadow.inherit = False
        
        if "visible" in shadow_config:
            shadow.visible = bool(shadow_config["visible"])
        
        if "color" in shadow_config:
            color = ColorFormatter.parse_color(shadow_config["color"])
            if color:
                shadow.shadow_type = shadow_config.get("type", "outer")
        
        # Additional shadow properties like blur, distance, angle could be added
