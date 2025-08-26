"""
JSON Generator for AI/RAG system
Converts content plans into pypptx-engine JSON specifications
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime


class JSONGenerator:
    """Generate pypptx-engine JSON from content plans"""
    
    def __init__(self):
        self.default_styles = {
            "title_font": {"name": "Helvetica Neue", "size": 56, "bold": True, "color": "#ffffff"},
            "subtitle_font": {"name": "Helvetica Neue", "size": 24, "color": "#a0a0a0"},
            "heading_font": {"name": "Segoe UI", "size": 36, "bold": True, "color": "#2c3e50"},
            "body_font": {"name": "Segoe UI", "size": 18, "color": "#2c3e50"},
            "bullet_font": {"name": "Segoe UI", "size": 18, "color": "#ffffff"}
        }
        
        self.color_schemes = {
            "professional": {"primary": "#2c3e50", "secondary": "#34495e", "accent": "#5B2C6F", "text": "#ffffff"},
            "modern": {"primary": "#000000", "secondary": "#1a1a1a", "accent": "#007aff", "text": "#ffffff"},
            "creative": {"primary": "#667eea", "secondary": "#764ba2", "accent": "#ffc000", "text": "#ffffff"}
        }
        
        # Professional slide templates based on examples
        self.slide_templates = {
            "title_slide": self._get_title_slide_template,
            "content_slide": self._get_content_slide_template,
            "bullet_slide": self._get_bullet_slide_template,
            "image_slide": self._get_image_slide_template,
            "comparison_slide": self._get_comparison_slide_template
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
    
    def _generate_slide_json(self, slide_info: Dict[str, Any], parsed_plan: Dict[str, Any]) -> Dict[str, Any]:
        """Generate JSON for a single slide"""
        color_scheme = self.color_schemes.get(parsed_plan.get("color_scheme", "professional"))
        
        if slide_info["number"] == 1:
            return self._generate_title_slide(slide_info, color_scheme, parsed_plan)
        else:
            return self._generate_content_slide(slide_info, color_scheme)
    
    def _generate_title_slide(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str], 
                            parsed_plan: Dict[str, Any]) -> Dict[str, Any]:
        """Generate title slide JSON"""
        return {
            "title": "Title Slide",
            "background": {
                "type": "gradient",
                "gradient": {
                    "type": "linear",
                    "angle": 45,
                    "stops": [
                        {"position": 0, "color": color_scheme["primary"]},
                        {"position": 1, "color": color_scheme["secondary"]}
                    ]
                }
            },
            "elements": [
                {
                    "type": "text",
                    "content": parsed_plan.get("title", "AI Generated Presentation"),
                    "position": {"x": 1, "y": 2.5, "width": 8, "height": 2},
                    "font": {
                        **self.default_styles["title_font"],
                        "color": "#ffffff"
                    },
                    "alignment": "center"
                },
                {
                    "type": "text",
                    "content": parsed_plan.get("goal", "Generated with pypptx-engine AI"),
                    "position": {"x": 1, "y": 4.8, "width": 8, "height": 1},
                    "font": {
                        **self.default_styles["subtitle_font"],
                        "color": "#ffffff"
                    },
                    "alignment": "center"
                }
                {
                    "type": "text",
                    "text": parsed_plan.get("goal", "Professional Presentation"),
                    "x": 1,
                    "y": 4.5,
                    "w": 14,
                    "h": 1,
                    "transparent": True,
                    "font": {
                        "name": "Helvetica Neue",
                        "size": 24,
                        "color": "#a0a0a0"
                    },
                    "paragraph": {
                        "alignment": "center"
                    }
                }
            ]
        }
    
    def _get_bullet_slide_template(self, slide_info: Dict[str, Any], color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Professional bullet slide template"""
        main_points = slide_info.get("main_points", [])
        bullet_points = ["• " + point for point in main_points[:6]]  # Limit to 6 points
        
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
        slide_json = {
            "title": slide_info["title"],
            "background": self._generate_background(slide_info.get("background", "gradient"), color_scheme),
            "elements": []
        }
        
        # Add title
        slide_json["elements"].append({
            "type": "text",
            "content": slide_info["title"],
            "position": {"x": 0.5, "y": 0.5, "width": 9, "height": 1},
            "font": self.default_styles["heading_font"],
            "alignment": "left"
        })
        
        # Add content based on type
        content_type = slide_info.get("content_type", "text")
        
        if content_type == "bullet_list" or len(slide_info.get("main_points", [])) > 1:
            slide_json["elements"].append(self._generate_bullet_list(slide_info["main_points"]))
        elif content_type == "text":
            slide_json["elements"].append(self._generate_text_content(slide_info["main_points"]))
        elif content_type == "image":
            slide_json["elements"].extend(self._generate_image_content(slide_info))
        elif content_type == "chart":
            slide_json["elements"].append(self._generate_chart_placeholder(slide_info))
        elif content_type == "table":
            slide_json["elements"].append(self._generate_table_placeholder(slide_info))
        elif content_type == "flowchart":
            slide_json["elements"].extend(self._generate_flowchart_placeholder(slide_info))
        
        return slide_json
    
    def _generate_background(self, bg_type: str, color_scheme: Dict[str, str]) -> Dict[str, Any]:
        """Generate background configuration"""
        if "gradient" in bg_type.lower():
            return {
                "type": "gradient",
                "gradient": {
                    "type": "linear",
                    "angle": 135,
                    "stops": [
                        {"position": 0, "color": "#ffffff"},
                        {"position": 1, "color": "#f8f9fa"}
                    ]
                }
            }
        elif "image" in bg_type.lower():
            return {
                "type": "image",
                "image": {
                    "url": "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=1920&h=1080&fit=crop",
                    "opacity": 0.1
                }
            }
        else:
            return {"type": "solid", "color": "#ffffff"}
    
    def _generate_bullet_list(self, points: List[str]) -> Dict[str, Any]:
        """Generate bullet list element"""
        return {
            "type": "bullet_list",
            "items": points[:6],  # Limit to 6 points for readability
            "position": {"x": 1, "y": 2, "width": 8, "height": 4},
            "font": self.default_styles["bullet_font"],
            "bullet_style": "•",
            "line_spacing": 1.2
        }
    
    def _generate_text_content(self, points: List[str]) -> Dict[str, Any]:
        """Generate text content element"""
        content = "\n\n".join(points) if points else "Content will be added here."
        
        return {
            "type": "text",
            "content": content,
            "position": {"x": 1, "y": 2, "width": 8, "height": 4},
            "font": self.default_styles["body_font"],
            "alignment": "left",
            "line_spacing": 1.3
        }
    
    def _generate_image_content(self, slide_info: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Generate image content elements"""
        elements = []
        
        # Add placeholder image
        elements.append({
            "type": "image",
            "url": "https://images.unsplash.com/photo-1560472354-b33ff0c44a43?w=800&h=600&fit=crop",
            "position": {"x": 5, "y": 2, "width": 4, "height": 3},
            "alt_text": f"Image for {slide_info['title']}"
        })
        
        # Add text content alongside
        if slide_info.get("main_points"):
            elements.append({
                "type": "text",
                "content": "\n".join(slide_info["main_points"]),
                "position": {"x": 0.5, "y": 2, "width": 4, "height": 3},
                "font": self.default_styles["body_font"],
                "alignment": "left"
            })
        
        return elements
    
    def _generate_chart_placeholder(self, slide_info: Dict[str, Any]) -> Dict[str, Any]:
        """Generate chart placeholder"""
        return {
            "type": "chart",
            "chart_type": "column",
            "data": {
                "categories": ["Category 1", "Category 2", "Category 3"],
                "series": [
                    {"name": "Series 1", "values": [10, 20, 15]}
                ]
            },
            "position": {"x": 2, "y": 2, "width": 6, "height": 4},
            "title": f"Chart for {slide_info['title']}"
        }
    
    def _generate_table_placeholder(self, slide_info: Dict[str, Any]) -> Dict[str, Any]:
        """Generate table placeholder"""
        return {
            "type": "table",
            "data": [
                ["Header 1", "Header 2", "Header 3"],
                ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
                ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
            ],
            "position": {"x": 1, "y": 2, "width": 8, "height": 3},
            "style": {
                "header_font": {**self.default_styles["body_font"], "bold": True},
                "cell_font": self.default_styles["body_font"],
                "border": True
            }
        }
    
    def _generate_flowchart_placeholder(self, slide_info: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Generate flowchart placeholder elements"""
        elements = []
        
        # Create simple flowchart with rectangles
        points = slide_info.get("main_points", ["Step 1", "Step 2", "Step 3"])
        
        for i, point in enumerate(points[:4]):  # Limit to 4 steps
            y_pos = 2 + (i * 1.2)
            
            # Add rectangle shape
            elements.append({
                "type": "autoshape",
                "shape": "rectangle",
                "position": {"x": 3, "y": y_pos, "width": 4, "height": 0.8},
                "text": point,
                "font": self.default_styles["body_font"],
                "fill": {"color": "#e7f3ff"},
                "border": {"color": "#1f4e79", "width": 1}
            })
            
            # Add arrow (except for last item)
            if i < len(points) - 1 and i < 3:
                elements.append({
                    "type": "autoshape",
                    "shape": "down_arrow",
                    "position": {"x": 4.8, "y": y_pos + 0.9, "width": 0.4, "height": 0.2},
                    "fill": {"color": "#1f4e79"}
                })
        
        return elements
