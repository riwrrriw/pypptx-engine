"""
CLI to generate a PowerPoint presentation from a JSON specification.

Usage (via Poetry):
  poetry run pypptx-engine --input src/examples/input.json --output from_json.pptx
"""
from __future__ import annotations

import argparse
import json
import os
from typing import Any, Dict

from .engine import PPTXEngine


def load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate PPTX from JSON spec.")
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="Path to input JSON file.",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="from_json.pptx",
        help="Path to output PPTX file.",
    )
    parser.add_argument(
        "--assets-base",
        default=None,
        help="Optional base directory for relative asset paths (defaults to the JSON file's directory)",
    )
    parser.add_argument(
        "--validate",
        "-v",
        action="store_true",
        help="Validate JSON file without generating presentation",
    )
    parser.add_argument(
        "--no-engine-test",
        action="store_true",
        help="Skip engine compatibility test during validation",
    )
    return parser.parse_args()


def main() -> None:
    """Main CLI entry point."""
    args = parse_args()
    
    try:
        # Validation mode
        if args.validate:
            from .validate import JSONValidator
            validator = JSONValidator()
            success = validator.validate_file(
                args.input,
                test_engine=not args.no_engine_test
            )
            validator.print_results()
            
            if success:
                print(f"\n🎉 {args.input} is valid and ready to use!")
                exit(0)
            else:
                print(f"\n💥 {args.input} has validation errors")
                exit(1)
        
        # Normal generation mode
        # Load configuration
        config = load_json(args.input)
        
        # Determine base directory for assets
        base_dir = args.assets_base or os.path.dirname(os.path.abspath(args.input))
        
        # Create presentation using the engine
        engine = PPTXEngine()
        presentation = engine.create_presentation(config, base_dir)
        
        # Save presentation
        presentation.save(args.output)
        print(f"✅ Presentation saved to: {args.output}")
        
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {e}")
        exit(1)
    except json.JSONDecodeError as e:
        print(f"❌ Error: Invalid JSON - {e}")
        exit(1)
    except Exception as e:
        print(f"❌ Error: {e}")
        exit(1)
