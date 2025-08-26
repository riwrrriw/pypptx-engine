"""
Diagnostic tools for pypptx-engine presentation generation
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, List
from pathlib import Path

from .engine import PPTXEngine


class PresentationDiagnostics:
    """Diagnostic tools to identify and fix presentation issues."""
    
    def __init__(self):
        self.engine = PPTXEngine()
        self.issues = []
    
    def diagnose_json(self, json_path: str) -> Dict[str, Any]:
        """Diagnose issues in JSON configuration file."""
        results = {
            "valid_json": False,
            "has_presentation": False,
            "slide_count": 0,
            "issues": [],
            "warnings": []
        }
        
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            results["valid_json"] = True
            
            # Check presentation structure
            if "presentation" in config:
                results["has_presentation"] = True
                pres_config = config["presentation"]
                
                # Check slides
                if "slides" in pres_config:
                    slides = pres_config["slides"]
                    results["slide_count"] = len(slides)
                    
                    # Validate each slide
                    for i, slide in enumerate(slides):
                        self._validate_slide(slide, i, results)
                else:
                    results["issues"].append("No 'slides' array found in presentation")
            else:
                results["issues"].append("No 'presentation' object found in JSON")
                
        except json.JSONDecodeError as e:
            results["issues"].append(f"Invalid JSON: {e}")
        except FileNotFoundError:
            results["issues"].append(f"File not found: {json_path}")
        except Exception as e:
            results["issues"].append(f"Unexpected error: {e}")
        
        return results
    
    def _validate_slide(self, slide: Dict[str, Any], slide_index: int, results: Dict[str, Any]) -> None:
        """Validate individual slide configuration."""
        slide_prefix = f"Slide {slide_index + 1}"
        
        # Check layout
        if "layout" not in slide:
            results["warnings"].append(f"{slide_prefix}: No layout specified, using default")
        
        # Check shapes
        if "shapes" in slide:
            shapes = slide["shapes"]
            for j, shape in enumerate(shapes):
                self._validate_shape(shape, slide_index, j, results)
    
    def _validate_shape(self, shape: Dict[str, Any], slide_index: int, shape_index: int, results: Dict[str, Any]) -> None:
        """Validate individual shape configuration."""
        shape_prefix = f"Slide {slide_index + 1}, Shape {shape_index + 1}"
        
        # Check required fields
        if "type" not in shape:
            results["issues"].append(f"{shape_prefix}: Missing 'type' field")
            return
        
        shape_type = shape["type"]
        
        # Validate coordinates
        for coord in ["x", "y", "w", "h"]:
            if coord not in shape:
                results["warnings"].append(f"{shape_prefix}: Missing '{coord}' coordinate, using default")
        
        # Type-specific validation
        if shape_type == "chart":
            self._validate_chart(shape, shape_prefix, results)
        elif shape_type == "image":
            self._validate_image(shape, shape_prefix, results)
        elif shape_type == "autoshape":
            self._validate_autoshape(shape, shape_prefix, results)
    
    def _validate_chart(self, shape: Dict[str, Any], prefix: str, results: Dict[str, Any]) -> None:
        """Validate chart configuration."""
        if "chartType" not in shape:
            results["warnings"].append(f"{prefix}: No chartType specified, using default")
        
        if "series" not in shape or not shape["series"]:
            results["issues"].append(f"{prefix}: Chart has no data series")
    
    def _validate_image(self, shape: Dict[str, Any], prefix: str, results: Dict[str, Any]) -> None:
        """Validate image configuration."""
        if "path" not in shape:
            results["issues"].append(f"{prefix}: Image has no 'path' specified")
        else:
            # Check if file exists (relative to common locations)
            image_path = shape["path"]
            if not os.path.isabs(image_path):
                # Check common locations
                possible_paths = [
                    image_path,
                    f"src/examples/{image_path}",
                    f"assets/{image_path}"
                ]
                found = any(os.path.exists(p) for p in possible_paths)
                if not found:
                    results["warnings"].append(f"{prefix}: Image file may not exist: {image_path}")
    
    def _validate_autoshape(self, shape: Dict[str, Any], prefix: str, results: Dict[str, Any]) -> None:
        """Validate autoshape configuration."""
        if "shape_type" not in shape:
            results["warnings"].append(f"{prefix}: No shape_type specified for autoshape")
    
    def generate_report(self, json_path: str) -> str:
        """Generate a comprehensive diagnostic report."""
        results = self.diagnose_json(json_path)
        
        report = f"# Presentation Diagnostic Report\n\n"
        report += f"**File**: {json_path}\n\n"
        
        # Status
        if results["valid_json"]:
            report += "✅ **JSON Format**: Valid\n"
        else:
            report += "❌ **JSON Format**: Invalid\n"
        
        if results["has_presentation"]:
            report += "✅ **Presentation Structure**: Valid\n"
        else:
            report += "❌ **Presentation Structure**: Missing\n"
        
        report += f"📊 **Slide Count**: {results['slide_count']}\n\n"
        
        # Issues
        if results["issues"]:
            report += "## ❌ Critical Issues\n\n"
            for issue in results["issues"]:
                report += f"- {issue}\n"
            report += "\n"
        
        # Warnings
        if results["warnings"]:
            report += "## ⚠️ Warnings\n\n"
            for warning in results["warnings"]:
                report += f"- {warning}\n"
            report += "\n"
        
        if not results["issues"] and not results["warnings"]:
            report += "## ✅ All Good!\n\nNo issues or warnings found.\n\n"
        
        # Recommendations
        report += "## 🔧 Recommendations\n\n"
        if results["issues"]:
            report += "1. Fix critical issues before generating presentation\n"
        if results["warnings"]:
            report += "2. Review warnings for potential improvements\n"
        report += "3. Test with simple configuration first\n"
        report += "4. Use `poetry run pypptx-engine --input file.json --output test.pptx`\n"
        
        return report
    
    def fix_common_issues(self, json_path: str, output_path: str = None) -> str:
        """Automatically fix common issues in JSON configuration."""
        if output_path is None:
            base = Path(json_path).stem
            output_path = f"{base}_fixed.json"
        
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            fixed_config = self._apply_fixes(config)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(fixed_config, f, indent=2)
            
            return f"✅ Fixed configuration saved to: {output_path}"
            
        except Exception as e:
            return f"❌ Error fixing configuration: {e}"
    
    def _apply_fixes(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Apply automatic fixes to configuration."""
        # Ensure presentation structure
        if "presentation" not in config:
            config = {"presentation": config}
        
        pres = config["presentation"]
        
        # Add default properties if missing
        if "properties" not in pres:
            pres["properties"] = {
                "title": "Generated Presentation",
                "author": "pypptx-engine"
            }
        
        # Add default size if missing
        if "size" not in pres:
            pres["size"] = {"width_in": 16, "height_in": 9}
        
        # Fix slides
        if "slides" in pres:
            for slide in pres["slides"]:
                self._fix_slide(slide)
        
        return config
    
    def _fix_slide(self, slide: Dict[str, Any]) -> None:
        """Fix common slide issues."""
        # Add default layout
        if "layout" not in slide:
            slide["layout"] = 6  # Blank layout
        
        # Fix shapes
        if "shapes" in slide:
            for shape in slide["shapes"]:
                self._fix_shape(shape)
    
    def _fix_shape(self, shape: Dict[str, Any]) -> None:
        """Fix common shape issues."""
        # Add default coordinates
        defaults = {"x": 1, "y": 1, "w": 4, "h": 1}
        for coord, default in defaults.items():
            if coord not in shape:
                shape[coord] = default
        
        # Fix chart issues
        if shape.get("type") == "chart":
            if "chartType" not in shape:
                shape["chartType"] = "COLUMN_CLUSTERED"
            if "categories" not in shape:
                shape["categories"] = ["A", "B", "C"]
            if "series" not in shape or not shape["series"]:
                shape["series"] = [{"name": "Data", "values": [1, 2, 3]}]


def main():
    """CLI interface for diagnostics."""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python -m pypptx_engine.diagnostics <json_file>")
        return
    
    json_file = sys.argv[1]
    diagnostics = PresentationDiagnostics()
    
    # Generate report
    report = diagnostics.generate_report(json_file)
    print(report)
    
    # Offer to fix issues
    results = diagnostics.diagnose_json(json_file)
    if results["issues"]:
        response = input("\nWould you like to generate a fixed version? (y/n): ")
        if response.lower() == 'y':
            fix_result = diagnostics.fix_common_issues(json_file)
            print(fix_result)


if __name__ == "__main__":
    main()
