"""
Enhanced JSON Generator for AI/RAG system
Converts content plans into professional pypptx-engine JSON specifications
Based on analysis of high-quality examples
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime


class EnhancedJSONGenerator:
    """Generate professional pypptx-engine JSON from content plans"""
    
    def __init__(self):
        # Professional styles based on examples
        self.default_styles = {
            "title_font": {"name": "Helvetica Neue", "size": 56, "bold": True, "color": "#ffffff"},
            "subtitle_font": {"name": "Helvetica Neue", "size": 24, "color": "#a0a0a0"},
            "heading_font": {"name": "Segoe UI", "size": 36, "bold": True, "color": "#2c3e50"},
            "body_font": {"name": "Segoe UI", "size": 18, "color": "#2c3e50"},
            "bullet_font": {"name": "Segoe UI", "size": 18, "color": "#ffffff"}
        }
        
        # Enhanced color schemes from examples
        self.color_schemes = {
            "professional": {"primary": "#2c3e50", "secondary": "#34495e", "accent": "#5B2C6F", "text": "#ffffff"},
            "modern": {"primary": "#000000", "secondary": "#1a1a1a", "accent": "#007aff", "text": "#ffffff"},
            "creative": {"primary": "#667eea", "secondary": "#764ba2", "accent": "#ffc000", "text": "#ffffff"}
        }
    
    def generate_from_content_plan(self, content_plan_path: Path) -> Dict[str, Any]:
        """Generate JSON specification from content plan"""
        with open(content_plan_path, 'r', encoding='utf-8') as f:
            content_plan = f.read()
        
        # Parse content plan
        parsed_plan = self._parse_content_plan(content_plan)
        
        # Generate JSON structure with professional format
        json_spec = {
            "presentation": {
                "properties": {
                    "title": parsed_plan.get("title", "AI Generated Presentation"),
                    "author": "pypptx-engine AI",
                    "subject": parsed_plan.get("goal", "Professional Presentation"),
                    "keywords": "ai-generated, professional, presentation"
                },
                "size": {"width_in": 16, "height_in": 9}
            },
            "slides": []
        }
        
        # Generate slides with enhanced templates
        for slide_info in parsed_plan.get("slides", []):
            slide_json = self._generate_enhanced_slide(slide_info, parsed_plan)
            json_spec["slides"].append(slide_json)
        
        return json_spec
    
    def _parse_content_plan(self, content_plan: str) -> Dict[str, Any]:
        """Parse content plan markdown into structured data"""
        parsed = {
            "title": "AI Generated Presentation",
            "slides": [],
            "design_preferences": {},
            "color_scheme": "professional"
        }
        
        # Extract project title
        title_match = re.search(r'\*\*Project Name\*\*:\s*(.+)', content_plan)
        if title_match:
            parsed["title"] = title_match.group(1).strip()
        
        # Extract presentation goal
        goal_match = re.search(r'\*\*Presentation Goal\*\*:\s*(.+)', content_plan)
        if goal_match:
            parsed["goal"] = goal_match.group(1).strip()
        
        # Extract color scheme
        color_match = re.search(r'\*\*Color Scheme\*\*:\s*(.+)', content_plan)
        if color_match:
            color_text = color_match.group(1).lower()
            if "modern" in color_text:
                parsed["color_scheme"] = "modern"
            elif "creative" in color_text:
                parsed["color_scheme"] = "creative"
        
        # Parse slides
        slide_sections = re.findall(r'### Slide (\d+): (.+?)\n(.*?)(?=### Slide|\## Design|$)', 
                                  content_plan, re.DOTALL)
        
        for slide_num, slide_title, slide_content in slide_sections:
            slide_info = self._parse_slide_section(slide_num, slide_title, slide_content)
            parsed["slides"].append(slide_info)
        
        return parsed
    
    def _parse_slide_section(self, slide_num: str, slide_title: str, slide_content: str) -> Dict[str, Any]:
        """Parse individual slide section"""
        slide_info = {
            "number": int(slide_num),
            "title": slide_title.strip(),
            "content_type": "text",
            "main_points": [],
            "background": "gradient",
            "source": ""
        }
        
        # Extract content type
        content_type_match = re.search(r'\*\*Content Type\*\*:\s*(.+)', slide_content)
        if content_type_match:
            slide_info["content_type"] = content_type_match.group(1).strip()
        
        # Extract main points
        points_section = re.search(r'\*\*Main Points\*\*:\s*\n(.*?)(?=\*\*|$)', slide_content, re.DOTALL)
        if points_section:
            points_text = points_section.group(1)
            points = re.findall(r'- (.+)', points_text)
            slide_info["main_points"] = [point.strip() for point in points if point.strip()]
        
        # Extract background
        bg_match = re.search(r'\*\*Background\*\*:\s*(.+)', slide_content)
        if bg_match:
            slide_info["background"] = bg_match.group(1).strip()
        
        # Extract source
        source_match = re.search(r'\*\*Source\*\*:\s*(.+)', slide_content)
        if source_match:
            slide_info["source"] = source_match.group(1).strip()
        
        return slide_info
    
    def _generate_enhanced_slide(self, slide_info: Dict[str, Any], parsed_plan: Dict[str, Any]) -> Dict[str, Any]:
        """Generate enhanced slide with professional templates"""
        color_scheme = self.color_schemes.get(parsed_plan.get("color_scheme", "professional"))
        
        if slide_info["number"] == 1:
            return self._get_title_slide_template(slide_info, color_scheme, parsed_plan)
        else:
            # Determine slide type based on content
            content_type = slide_info.get("content_type", "text")
            main_points = slide_info.get("main_points", [])
            
            if content_type == "bullet_list" or len(main_points) > 3:
                return self._get_bullet_slide_template(slide_info, color_scheme)
            elif "image" in content_type.lower() or "visual" in slide_info.get("title", "").lower():
                return self._get_image_slide_template(slide_info, color_scheme)
            elif "comparison" in slide_info.get("title", "").lower() or "vs" in slide_info.get("title", "").lower():
                return self._get_comparison_slide_template(slide_info, color_scheme)
            else:
                return self._get_content_slide_template(slide_info, color_scheme)
    
    def _get_title_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str], parsed_plan: Dict[str, Any]) -> Dict[str, Any]:
        """Professional title slide template based on iPhone 16 Pro example"""
        return {
            "layout": 6,
            "background": {
                "type": "gradient",
                "colors": [color_scheme["primary"], color_scheme["secondary"]],
                "direction": "radial"
            },
            "transition": {
                "type": "fade",
                "duration": 1.0
            },
            "shapes": [
                {
                    "type": "text",
                    "text": "Introducing",
                    "x": 1,
                    "y": 2,
                    "w": 7,
                    "h": 1,
                    "transparent": True,
                    "font": {
                        "name": "Helvetica Neue",
                        "size": 32,
                        "color": "#ffffff"
                    },
                    "paragraph": {
                        "alignment": "left"
                    }
                },
                {
                    "type": "text",
                    "text": parsed_plan.get("title", "AI Generated Presentation"),
                    "x": 1,
                    "y": 3,
                    "w": 7,
                    "h": 1.5,
                    "transparent": True,
                    "font": {
                        "name": "Helvetica Neue",
                        "size": 56,
                        "bold": True,
                        "color": "#ffffff"
                    },
                    "paragraph": {
                        "alignment": "left"
                    }
                },
                {
                    "type": "text",
                    "text": parsed_plan.get("goal", "Professional AI-Generated Presentation"),
                    "x": 1,
                    "y": 4.8,
                    "w": 7,
                    "h": 0.8,
                    "transparent": True,
                    "font": {
                        "name": "Helvetica Neue",
                        "size": 24,
                        "color": "#a0a0a0"
                    },
                    "paragraph": {
                        "alignment": "left"
                    }
                },
                {
                    "type": "image",
                    "url": "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&q=80",
                    "x": 8.5,
                    "y": 1.5,
                    "w": 6,
                    "h": 6
                }
            ]
        }
    
    def _get_bullet_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Professional bullet slide template based on CardX example"""
        main_points = slide_info.get("main_points", [])
        bullet_points = ["✓ " + point for point in main_points[:6]]  # Limit to 6 points
        
        return {
            "layout": 6,
            "background": {
                "type": "gradient",
                "colors": [color_scheme["primary"], color_scheme["secondary"]],
                "direction": "vertical"
            },
            "shapes": [
                {
                    "type": "text",
                    "text": slide_info["title"],
                    "x": 1,
                    "y": 0.5,
                    "w": 14,
                    "h": 1,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 36,
                        "bold": True,
                        "color": "#ffffff"
                    },
                    "paragraph": {
                        "alignment": "center"
                    }
                },
                {
                    "type": "autoshape",
                    "shape_type": "RECTANGLE",
                    "x": 1.5,
                    "y": 2,
                    "w": 13,
                    "h": 5,
                    "fill": {"type": "solid", "color": "#ffffff"},
                    "line": {"color": "#e0e0e0", "width": 1},
                    "shadow": {"visible": True, "color": "#00000020", "blur": 10, "distance": 3}
                },
                {
                    "type": "text",
                    "text": bullet_points,
                    "x": 2,
                    "y": 2.5,
                    "w": 12,
                    "h": 4,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 18,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": 1.4
                    }
                }
            ]
        }
    
    def _get_content_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Professional content slide template"""
        main_points = slide_info.get("main_points", [])
        content_text = "\n".join(main_points) if main_points else "Content will be generated based on your source materials."
        
        return {
            "layout": 6,
            "background": "#ffffff",
            "shapes": [
                {
                    "type": "text",
                    "text": slide_info["title"],
                    "x": 1,
                    "y": 0.5,
                    "w": 14,
                    "h": 1,
                    "font": {
                        "name": "Segoe UI",
                        "size": 36,
                        "bold": True,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "center"
                    }
                },
                {
                    "type": "text",
                    "text": content_text,
                    "x": 2,
                    "y": 2,
                    "w": 12,
                    "h": 5,
                    "font": {
                        "name": "Segoe UI",
                        "size": 18,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": 1.2
                    },
                    "fill": {
                        "type": "solid",
                        "color": "#f8f9fa"
                    },
                    "line": {
                        "color": "#e0e0e0",
                        "width": 1
                    }
                }
            ]
        }
    
    def _get_image_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Professional image slide template"""
        return {
            "layout": 6,
            "background": {
                "type": "gradient",
                "colors": [color_scheme["primary"], color_scheme["secondary"]],
                "direction": "diagonal"
            },
            "shapes": [
                {
                    "type": "text",
                    "text": slide_info["title"],
                    "x": 1,
                    "y": 0.5,
                    "w": 14,
                    "h": 1,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 36,
                        "bold": True,
                        "color": "#ffffff"
                    },
                    "paragraph": {
                        "alignment": "center"
                    }
                },
                {
                    "type": "image",
                    "url": "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=800&q=80",
                    "x": 2,
                    "y": 2,
                    "w": 6,
                    "h": 4
                },
                {
                    "type": "autoshape",
                    "shape_type": "RECTANGLE",
                    "x": 9,
                    "y": 2,
                    "w": 6,
                    "h": 4,
                    "fill": {"type": "solid", "color": "#ffffff"},
                    "line": {"color": "#e0e0e0", "width": 1},
                    "shadow": {"visible": True, "color": "#00000020", "blur": 8, "distance": 2}
                },
                {
                    "type": "text",
                    "text": "\n".join(slide_info.get("main_points", ["Visual content enhances understanding", "Professional design creates impact", "Data-driven insights"])),
                    "x": 9.5,
                    "y": 2.5,
                    "w": 5,
                    "h": 3,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 16,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": 1.4
                    }
                }
            ]
        }
    
    def _get_comparison_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Professional comparison slide template"""
        main_points = slide_info.get("main_points", [])
        left_points = main_points[:len(main_points)//2] if main_points else ["Option A", "Feature 1", "Benefit 1"]
        right_points = main_points[len(main_points)//2:] if main_points else ["Option B", "Feature 2", "Benefit 2"]
        
        return {
            "layout": 6,
            "background": "#f8f9fa",
            "shapes": [
                {
                    "type": "text",
                    "text": slide_info["title"],
                    "x": 1,
                    "y": 0.5,
                    "w": 14,
                    "h": 1,
                    "font": {
                        "name": "Segoe UI",
                        "size": 36,
                        "bold": True,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "center"
                    }
                },
                {
                    "type": "autoshape",
                    "shape_type": "RECTANGLE",
                    "x": 1,
                    "y": 2,
                    "w": 6.5,
                    "h": 5,
                    "fill": {"type": "solid", "color": "#ffffff"},
                    "line": {"color": color_scheme["primary"], "width": 2},
                    "shadow": {"visible": True, "color": "#00000015", "blur": 8, "distance": 2}
                },
                {
                    "type": "autoshape",
                    "shape_type": "RECTANGLE",
                    "x": 8.5,
                    "y": 2,
                    "w": 6.5,
                    "h": 5,
                    "fill": {"type": "solid", "color": "#ffffff"},
                    "line": {"color": color_scheme["accent"], "width": 2},
                    "shadow": {"visible": True, "color": "#00000015", "blur": 8, "distance": 2}
                },
                {
                    "type": "text",
                    "text": "\n".join(left_points),
                    "x": 1.5,
                    "y": 2.5,
                    "w": 5.5,
                    "h": 4,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 16,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": 1.4
                    }
                },
                {
                    "type": "text",
                    "text": "\n".join(right_points),
                    "x": 9,
                    "y": 2.5,
                    "w": 5.5,
                    "h": 4,
                    "transparent": True,
                    "font": {
                        "name": "Segoe UI",
                        "size": 16,
                        "color": "#2c3e50"
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": 1.4
                    }
                }
            ]
        }


def save_json_spec(json_spec: Dict[str, Any], output_path: Path) -> None:
    """Save JSON specification to file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(json_spec, f, indent=2, ensure_ascii=False)
