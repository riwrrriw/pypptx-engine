#!/usr/bin/env python3
"""
Test script to verify connector fix
"""
import sys
import os
import json

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from pypptx_engine import PPTXEngine

def test_connector_fix():
    """Test the fixed connector implementation."""
    
    # Simple test with the problematic connector configuration
    test_config = {
        "presentation": {
            "properties": {
                "title": "Connector Fix Test",
                "author": "PPTX Engine"
            },
            "slides": [
                {
                    "layout": 6,
                    "background": "#2c3e50",
                    "shapes": [
                        {
                            "type": "text",
                            "text": "Connector Test",
                            "x": 1,
                            "y": 0.5,
                            "w": 14,
                            "h": 1,
                            "font": {
                                "size": 32,
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
                            "text": "Start",
                            "x": 2,
                            "y": 2,
                            "w": 3,
                            "h": 2,
                            "fill": {
                                "type": "solid",
                                "color": "#e74c3c"
                            },
                            "line": {
                                "color": "#c0392b",
                                "width": 2
                            },
                            "font": {
                                "color": "#ffffff",
                                "bold": True
                            }
                        },
                        {
                            "type": "autoshape",
                            "shape_type": "RECTANGLE",
                            "text": "End",
                            "x": 10,
                            "y": 2,
                            "w": 3,
                            "h": 2,
                            "fill": {
                                "type": "solid",
                                "color": "#27ae60"
                            },
                            "line": {
                                "color": "#229954",
                                "width": 2
                            },
                            "font": {
                                "color": "#ffffff",
                                "bold": True
                            }
                        },
                        {
                            "type": "connector",
                            "connector_type": "STRAIGHT",
                            "x": 2.5,
                            "y": 4.5,
                            "w": 11,
                            "h": 0.5,
                            "line": {
                                "color": "#ffffff",
                                "width": 3
                            }
                        }
                    ]
                }
            ]
        }
    }
    
    # Create the engine
    engine = PPTXEngine()
    
    # Generate the presentation
    output_path = os.path.join(os.path.dirname(__file__), '..', 'output', 'connector_test.pptx')
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    try:
        engine.create_presentation(test_config, output_path)
        print(f"✅ Connector test passed: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Connector test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🔄 Testing connector fix...")
    success = test_connector_fix()
    
    if success:
        print("🎉 Connector fix verified!")
    else:
        print("⚠️  Connector fix needs more work.")
