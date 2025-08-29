#!/usr/bin/env python3
"""
JSON Validation CLI for pypptx-engine
Validates JSON configuration files and tests compatibility with the engine
"""
import argparse
import json
import sys
import traceback
from pathlib import Path
from typing import Dict, Any, List

from .engine import PPTXEngine
from .formatters import ColorFormatter


class JSONValidator:
    """Validates JSON configuration files for pypptx-engine."""
    
    def __init__(self):
        self.engine = PPTXEngine()
        self.errors = []
        self.warnings = []
    
    def validate_json_syntax(self, file_path: str) -> bool:
        """Validate JSON syntax."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                json.load(f)
            return True
        except json.JSONDecodeError as e:
            self.errors.append(f"JSON Syntax Error: {e}")
            return False
        except Exception as e:
            self.errors.append(f"File Error: {e}")
            return False
    
    def validate_structure(self, config: Dict[str, Any]) -> bool:
        """Validate the overall structure of the configuration."""
        valid = True
        
        # Check for presentation wrapper
        if "presentation" not in config:
            self.errors.append("Missing 'presentation' root object")
            valid = False
            return valid
        
        presentation = config["presentation"]
        
        # Check for required fields
        if "slides" not in presentation:
            self.errors.append("Missing 'slides' array in presentation")
            valid = False
        
        # Validate slides
        if "slides" in presentation:
            slides = presentation["slides"]
            if not isinstance(slides, list):
                self.errors.append("'slides' must be an array")
                valid = False
            elif len(slides) == 0:
                self.warnings.append("Presentation has no slides")
        
        return valid
    
    def validate_slides(self, slides: List[Dict[str, Any]]) -> bool:
        """Validate individual slides."""
        valid = True
        
        for i, slide in enumerate(slides):
            slide_errors = []
            
            # Validate layout
            layout = slide.get("layout")
            if layout is not None and not isinstance(layout, int):
                slide_errors.append("'layout' must be an integer")
            
            # Validate shapes
            shapes = slide.get("shapes", [])
            if not isinstance(shapes, list):
                slide_errors.append("'shapes' must be an array")
            else:
                for j, shape in enumerate(shapes):
                    shape_errors = self.validate_shape(shape)
                    if shape_errors:
                        slide_errors.extend([f"Shape {j}: {err}" for err in shape_errors])
            
            # Validate background
            background = slide.get("background")
            if background:
                bg_errors = self.validate_background(background)
                if bg_errors:
                    slide_errors.extend([f"Background: {err}" for err in bg_errors])
            
            if slide_errors:
                self.errors.extend([f"Slide {i}: {err}" for err in slide_errors])
                valid = False
        
        return valid
    
    def validate_shape(self, shape: Dict[str, Any]) -> List[str]:
        """Validate individual shape configuration."""
        errors = []
        
        # Check required type
        if "type" not in shape:
            errors.append("Missing 'type' field")
            return errors
        
        shape_type = shape["type"]
        
        # Validate coordinates
        for coord in ["x", "y", "w", "h"]:
            if coord in shape and not isinstance(shape[coord], (int, float)):
                errors.append(f"'{coord}' must be a number")
        
        # Type-specific validation
        if shape_type == "text":
            if "text" not in shape:
                errors.append("Text shape missing 'text' field")
        
        elif shape_type == "table":
            errors.extend(self.validate_table_shape(shape))
        
        elif shape_type == "image":
            if "url" not in shape and "image_path" not in shape:
                errors.append("Image shape missing 'url' or 'image_path' field")
        
        elif shape_type == "chart":
            if "chart_type" not in shape:
                errors.append("Chart shape missing 'chart_type' field")
        
        elif shape_type == "flowchart":
            if "elements" not in shape:
                errors.append("Flowchart shape missing 'elements' field")
        
        return errors
    
    def validate_table_shape(self, shape: Dict[str, Any]) -> List[str]:
        """Validate table-specific configuration."""
        errors = []
        
        # Check rows and cols
        for field in ["rows", "cols"]:
            if field in shape and not isinstance(shape[field], int):
                errors.append(f"'{field}' must be an integer")
        
        # Validate data structure
        data = shape.get("data", [])
        if data and not isinstance(data, list):
            errors.append("'data' must be an array")
        
        # Validate merged cells
        merged_cells = shape.get("merged_cells", [])
        if merged_cells and not isinstance(merged_cells, list):
            errors.append("'merged_cells' must be an array")
        else:
            for i, merge in enumerate(merged_cells):
                if not isinstance(merge, dict):
                    errors.append(f"merged_cells[{i}] must be an object")
                    continue
                
                required_fields = ["start_row", "start_col", "end_row", "end_col"]
                for field in required_fields:
                    if field not in merge:
                        errors.append(f"merged_cells[{i}] missing '{field}'")
                    elif not isinstance(merge[field], int):
                        errors.append(f"merged_cells[{i}].{field} must be an integer")
        
        return errors
    
    def validate_background(self, background: Dict[str, Any]) -> List[str]:
        """Validate background configuration."""
        errors = []
        
        if not isinstance(background, dict):
            return ["Background must be an object"]
        
        bg_type = background.get("type", "solid")
        
        if bg_type == "image":
            if "url" not in background and "image_path" not in background:
                errors.append("Image background missing 'url' or 'image_path'")
        
        elif bg_type == "solid":
            if "color" not in background:
                errors.append("Solid background missing 'color'")
        
        elif bg_type == "gradient":
            if "colors" not in background:
                errors.append("Gradient background missing 'colors' array")
            elif not isinstance(background["colors"], list):
                errors.append("Gradient 'colors' must be an array")
        
        return errors
    
    def test_engine_compatibility(self, file_path: str) -> bool:
        """Test if the JSON can be processed by the engine."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # Try to create presentation
            prs = self.engine.create_presentation(config, str(Path(file_path).parent))
            
            # Try to save to a temporary location
            temp_output = Path(file_path).parent / "temp_validation.pptx"
            prs.save(str(temp_output))
            
            # Clean up
            if temp_output.exists():
                temp_output.unlink()
            
            return True
            
        except Exception as e:
            self.errors.append(f"Engine Compatibility Error: {e}")
            return False
    
    def validate_file(self, file_path: str, test_engine: bool = True) -> bool:
        """Validate a complete JSON file."""
        self.errors = []
        self.warnings = []
        
        print(f"Validating: {file_path}")
        print("=" * 50)
        
        # Step 1: JSON syntax
        if not self.validate_json_syntax(file_path):
            return False
        print("✅ JSON syntax valid")
        
        # Step 2: Load and validate structure
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            self.errors.append(f"Failed to load JSON: {e}")
            return False
        
        if not self.validate_structure(config):
            return False
        print("✅ Structure valid")
        
        # Step 3: Validate slides
        presentation = config["presentation"]
        slides = presentation.get("slides", [])
        if not self.validate_slides(slides):
            return False
        print("✅ Slides valid")
        
        # Step 4: Test engine compatibility
        if test_engine:
            if not self.test_engine_compatibility(file_path):
                return False
            print("✅ Engine compatibility confirmed")
        
        return True
    
    def print_results(self):
        """Print validation results."""
        if self.warnings:
            print("\n⚠️  Warnings:")
            for warning in self.warnings:
                print(f"  - {warning}")
        
        if self.errors:
            print("\n❌ Errors:")
            for error in self.errors:
                print(f"  - {error}")
            return False
        else:
            print("\n✅ All validations passed!")
            return True


def main():
    """Main CLI function."""
    parser = argparse.ArgumentParser(
        description="Validate JSON configuration files for pypptx-engine"
    )
    parser.add_argument(
        "file",
        help="JSON file to validate"
    )
    parser.add_argument(
        "--no-engine-test",
        action="store_true",
        help="Skip engine compatibility test"
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Verbose output"
    )
    
    args = parser.parse_args()
    
    if not Path(args.file).exists():
        print(f"❌ File not found: {args.file}")
        sys.exit(1)
    
    validator = JSONValidator()
    
    try:
        success = validator.validate_file(
            args.file,
            test_engine=not args.no_engine_test
        )
        
        validator.print_results()
        
        if success:
            print(f"\n🎉 {args.file} is valid and ready to use!")
            sys.exit(0)
        else:
            print(f"\n💥 {args.file} has validation errors")
            sys.exit(1)
            
    except Exception as e:
        print(f"❌ Validation failed with error: {e}")
        if args.verbose:
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
