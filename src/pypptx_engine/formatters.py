"""
Formatting utilities for colors, fonts, fills, and other visual properties
"""
from __future__ import annotations

from typing import Any, Dict, Optional, Union

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR, MSO_FILL_TYPE, MSO_PATTERN_TYPE, MSO_LINE_DASH_STYLE
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT, MSO_TEXT_UNDERLINE_TYPE
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
            pattern_type = fill_config.get("pattern_type", "PERCENT_5")
            if hasattr(MSO_PATTERN_TYPE, pattern_type):
                fill.pattern_type = getattr(MSO_PATTERN_TYPE, pattern_type)
            
            # Set pattern colors
            fore_color = ColorFormatter.parse_color(fill_config.get("fore_color"))
            back_color = ColorFormatter.parse_color(fill_config.get("back_color"))
            if fore_color:
                fill.fore_color.rgb = fore_color
            if back_color:
                fill.back_color.rgb = back_color
        elif fill_type == "picture":
            # Picture fill would need image path
            pass
        elif fill_type == "none":
            fill.background()


class FontFormatter:
    """Handle font formatting and text properties."""
    
    @staticmethod
    def apply_font_formatting(font, font_config: Dict[str, Any]) -> None:
        """Apply font formatting from configuration with enhanced options."""
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
            underline_type = font_config["underline"]
            if isinstance(underline_type, bool):
                font.underline = underline_type
            elif isinstance(underline_type, str):
                underline_type = underline_type.upper()
                if hasattr(MSO_TEXT_UNDERLINE_TYPE, underline_type):
                    font.underline = getattr(MSO_TEXT_UNDERLINE_TYPE, underline_type)
        
        # Enhanced color support
        if "color" in font_config:
            color_spec = font_config["color"]
            if isinstance(color_spec, dict):
                # Advanced color with theme support
                if "theme" in color_spec:
                    theme_color = color_spec["theme"].upper()
                    if hasattr(MSO_THEME_COLOR, theme_color):
                        font.color.theme_color = getattr(MSO_THEME_COLOR, theme_color)
                elif "rgb" in color_spec or "hex" in color_spec:
                    color = ColorFormatter.parse_color(color_spec)
                    if color:
                        font.color.rgb = color
            else:
                # Simple color string
                color = ColorFormatter.parse_color(color_spec)
                if color:
                    font.color.rgb = color
        
        # Additional font properties
        if "strikethrough" in font_config:
            # Note: python-pptx doesn't directly support strikethrough
            pass
        
        if "superscript" in font_config:
            if font_config["superscript"]:
                font.superscript = True
        
        if "subscript" in font_config:
            if font_config["subscript"]:
                font.subscript = True
    
    @staticmethod
    def apply_paragraph_formatting(paragraph, para_config: Dict[str, Any]) -> None:
        """Apply paragraph-level formatting with enhanced options."""
        if not para_config:
            return
        
        # Text alignment
        alignment = para_config.get("alignment", "").lower()
        if alignment == "left":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        elif alignment == "center":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        elif alignment == "right":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.RIGHT
        elif alignment == "justify":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.JUSTIFY
        elif alignment == "distribute":
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.DISTRIBUTE
        
        # Spacing controls
        if "space_before" in para_config:
            paragraph.space_before = Pt(para_config["space_before"])
        
        if "space_after" in para_config:
            paragraph.space_after = Pt(para_config["space_after"])
        
        if "line_spacing" in para_config:
            line_spacing = para_config["line_spacing"]
            if isinstance(line_spacing, (int, float)):
                paragraph.line_spacing = line_spacing
            elif isinstance(line_spacing, dict):
                # Advanced line spacing with units
                spacing_value = line_spacing.get("value", 1.0)
                spacing_unit = line_spacing.get("unit", "multiple")  # "multiple", "points", "lines"
                if spacing_unit == "multiple":
                    paragraph.line_spacing = spacing_value
                elif spacing_unit == "points":
                    paragraph.line_spacing = Pt(spacing_value)
        
        # Indentation
        if "left_indent" in para_config:
            paragraph.left_indent = Inches(para_config["left_indent"])
        
        if "right_indent" in para_config:
            paragraph.right_indent = Inches(para_config["right_indent"])
        
        if "first_line_indent" in para_config:
            paragraph.first_line_indent = Inches(para_config["first_line_indent"])
        
        # Bullet/numbering level
        if "level" in para_config:
            paragraph.level = para_config["level"]


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
        
        if "dash_style" in line_config:
            dash_style = line_config["dash_style"].upper()
            if hasattr(MSO_LINE_DASH_STYLE, dash_style):
                line.dash_style = getattr(MSO_LINE_DASH_STYLE, dash_style)


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
