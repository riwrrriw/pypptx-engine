"""
Template system for PPTX Engine.
Provides predefined templates and theme management.
"""

from typing import Dict, Any, List, Optional
import json
import os
from pathlib import Path


class TemplateManager:
    """Manages presentation templates and themes."""
    
    def __init__(self):
        self.templates = {}
        self.themes = {}
        self._load_built_in_templates()
        self._load_built_in_themes()
    
    def _load_built_in_templates(self):
        """Load built-in presentation templates."""
        self.templates = {
            "corporate": {
                "name": "Corporate",
                "description": "Professional corporate presentation template",
                "default_layout": 6,
                "slide_defaults": {
                    "background": {
                        "type": "gradient",
                        "direction": "vertical",
                        "colors": ["#f8f9fa", "#e9ecef"]
                    },
                    "title_style": {
                        "font": {
                            "name": "Calibri",
                            "size": 44,
                            "bold": True,
                            "color": "#2c3e50"
                        },
                        "position": {"x": 1, "y": 0.5, "w": 14, "h": 1.5}
                    },
                    "content_style": {
                        "font": {
                            "name": "Calibri",
                            "size": 24,
                            "color": "#34495e"
                        },
                        "position": {"x": 1, "y": 2.5, "w": 14, "h": 5}
                    }
                }
            },
            "modern": {
                "name": "Modern",
                "description": "Clean modern design template",
                "default_layout": 6,
                "slide_defaults": {
                    "background": {
                        "type": "solid",
                        "color": "#ffffff"
                    },
                    "title_style": {
                        "font": {
                            "name": "Segoe UI",
                            "size": 48,
                            "bold": True,
                            "color": "#0078d4"
                        },
                        "position": {"x": 1, "y": 1, "w": 14, "h": 1.5}
                    },
                    "content_style": {
                        "font": {
                            "name": "Segoe UI",
                            "size": 20,
                            "color": "#323130"
                        },
                        "position": {"x": 1, "y": 3, "w": 14, "h": 5}
                    }
                }
            },
            "creative": {
                "name": "Creative",
                "description": "Vibrant creative presentation template",
                "default_layout": 6,
                "slide_defaults": {
                    "background": {
                        "type": "gradient",
                        "direction": "diagonal",
                        "colors": ["#667eea", "#764ba2"]
                    },
                    "title_style": {
                        "font": {
                            "name": "Arial",
                            "size": 52,
                            "bold": True,
                            "color": "#ffffff"
                        },
                        "position": {"x": 1, "y": 1, "w": 14, "h": 2},
                        "shadow": {
                            "visible": True,
                            "color": "#000000",
                            "blur": 8
                        }
                    },
                    "content_style": {
                        "font": {
                            "name": "Arial",
                            "size": 22,
                            "color": "#f8f9fa"
                        },
                        "position": {"x": 1, "y": 3.5, "w": 14, "h": 4.5}
                    }
                }
            },
            "academic": {
                "name": "Academic",
                "description": "Professional academic presentation template",
                "default_layout": 6,
                "slide_defaults": {
                    "background": {
                        "type": "solid",
                        "color": "#fefefe"
                    },
                    "title_style": {
                        "font": {
                            "name": "Times New Roman",
                            "size": 40,
                            "bold": True,
                            "color": "#1a365d"
                        },
                        "position": {"x": 1, "y": 0.8, "w": 14, "h": 1.5}
                    },
                    "content_style": {
                        "font": {
                            "name": "Times New Roman",
                            "size": 18,
                            "color": "#2d3748"
                        },
                        "position": {"x": 1, "y": 2.8, "w": 14, "h": 5}
                    }
                }
            }
        }
    
    def _load_built_in_themes(self):
        """Load built-in color themes."""
        self.themes = {
            "blue": {
                "name": "Blue Theme",
                "primary": "#0078d4",
                "secondary": "#106ebe",
                "accent": "#40e0d0",
                "background": "#f8f9fa",
                "text": "#323130",
                "text_light": "#605e5c"
            },
            "green": {
                "name": "Green Theme",
                "primary": "#107c10",
                "secondary": "#0b5394",
                "accent": "#00bcf2",
                "background": "#f3f2f1",
                "text": "#323130",
                "text_light": "#605e5c"
            },
            "purple": {
                "name": "Purple Theme",
                "primary": "#5c2d91",
                "secondary": "#8764b8",
                "accent": "#c239b3",
                "background": "#faf9f8",
                "text": "#323130",
                "text_light": "#605e5c"
            },
            "orange": {
                "name": "Orange Theme",
                "primary": "#d83b01",
                "secondary": "#ff8c00",
                "accent": "#ffb900",
                "background": "#fdf6e3",
                "text": "#323130",
                "text_light": "#605e5c"
            }
        }
    
    def get_template(self, template_name: str) -> Optional[Dict[str, Any]]:
        """Get a template by name."""
        return self.templates.get(template_name)
    
    def get_theme(self, theme_name: str) -> Optional[Dict[str, Any]]:
        """Get a theme by name."""
        return self.themes.get(theme_name)
    
    def apply_template_to_config(self, config: Dict[str, Any], template_name: str, theme_name: Optional[str] = None) -> Dict[str, Any]:
        """Apply template and theme to presentation configuration."""
        template = self.get_template(template_name)
        if not template:
            raise ValueError(f"Template '{template_name}' not found")
        
        theme = None
        if theme_name:
            theme = self.get_theme(theme_name)
            if not theme:
                raise ValueError(f"Theme '{theme_name}' not found")
        
        # Apply template to each slide
        for slide in config.get("presentation", {}).get("slides", []):
            self._apply_template_to_slide(slide, template, theme)
        
        return config
    
    def _apply_template_to_slide(self, slide: Dict[str, Any], template: Dict[str, Any], theme: Optional[Dict[str, Any]]):
        """Apply template styling to a single slide."""
        slide_defaults = template.get("slide_defaults", {})
        
        # Apply background if not specified
        if "background" not in slide and "background" in slide_defaults:
            slide["background"] = slide_defaults["background"].copy()
            if theme:
                self._apply_theme_to_background(slide["background"], theme)
        
        # Apply template styling to shapes
        for shape in slide.get("shapes", []):
            if shape.get("type") == "text":
                self._apply_template_to_text_shape(shape, slide_defaults, theme)
    
    def _apply_template_to_text_shape(self, shape: Dict[str, Any], slide_defaults: Dict[str, Any], theme: Optional[Dict[str, Any]]):
        """Apply template styling to text shapes."""
        # Determine if this is a title or content based on position/size
        is_title = shape.get("y", 0) < 2 and shape.get("h", 0) <= 2
        
        style_key = "title_style" if is_title else "content_style"
        default_style = slide_defaults.get(style_key, {})
        
        # Apply font styling if not specified
        if "font" not in shape and "font" in default_style:
            shape["font"] = default_style["font"].copy()
            if theme:
                self._apply_theme_to_font(shape["font"], theme, is_title)
        
        # Apply positioning if not fully specified
        if "position" in default_style:
            pos = default_style["position"]
            for key in ["x", "y", "w", "h"]:
                if key not in shape and key in pos:
                    shape[key] = pos[key]
        
        # Apply shadow if specified in template
        if "shadow" in default_style and "shadow" not in shape:
            shape["shadow"] = default_style["shadow"].copy()
    
    def _apply_theme_to_background(self, background: Dict[str, Any], theme: Dict[str, Any]):
        """Apply theme colors to background."""
        if background.get("type") == "solid":
            if "color" not in background:
                background["color"] = theme["background"]
        elif background.get("type") == "gradient":
            if "colors" not in background:
                background["colors"] = [theme["background"], theme["primary"]]
    
    def _apply_theme_to_font(self, font: Dict[str, Any], theme: Dict[str, Any], is_title: bool):
        """Apply theme colors to font."""
        if "color" not in font:
            font["color"] = theme["primary"] if is_title else theme["text"]
    
    def list_templates(self) -> List[str]:
        """Get list of available template names."""
        return list(self.templates.keys())
    
    def list_themes(self) -> List[str]:
        """Get list of available theme names."""
        return list(self.themes.keys())
    
    def create_template_config(self, template_name: str, theme_name: Optional[str] = None, 
                             title: str = "Sample Presentation", 
                             slides_content: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
        """Create a complete presentation configuration using template and theme."""
        template = self.get_template(template_name)
        if not template:
            raise ValueError(f"Template '{template_name}' not found")
        
        # Default slide content if none provided
        if not slides_content:
            slides_content = [
                {
                    "title": "Welcome",
                    "content": ["Introduction to our presentation", "Key topics we'll cover", "Questions and discussion"]
                },
                {
                    "title": "Overview",
                    "content": ["Main objectives", "Methodology", "Expected outcomes"]
                }
            ]
        
        config = {
            "presentation": {
                "properties": {
                    "title": title,
                    "author": "PPTX Engine",
                    "template": template_name,
                    "theme": theme_name
                },
                "size": {"width_in": 16, "height_in": 9},
                "slides": []
            }
        }
        
        # Create slides from content
        for slide_content in slides_content:
            slide = {
                "layout": template.get("default_layout", 6),
                "shapes": []
            }
            
            # Add title if provided
            if "title" in slide_content:
                slide["shapes"].append({
                    "type": "text",
                    "text": slide_content["title"]
                })
            
            # Add content if provided
            if "content" in slide_content:
                slide["shapes"].append({
                    "type": "text",
                    "text": slide_content["content"]
                })
            
            config["presentation"]["slides"].append(slide)
        
        # Apply template and theme
        return self.apply_template_to_config(config, template_name, theme_name)
