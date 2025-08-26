#!/usr/bin/env python3
"""
Test script to demonstrate flowchart functionality
"""
import sys
import os
import json

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from pypptx_engine import PPTXEngine

def test_flowchart():
    """Test flowchart creation with the pypptx-engine."""
    
    # Load the flowchart example
    example_path = os.path.join(os.path.dirname(__file__), 'flowchart_example.json')
    
    with open(example_path, 'r') as f:
        config = json.load(f)
    
    # Create the engine
    engine = PPTXEngine()
    
    # Generate the presentation
    output_path = os.path.join(os.path.dirname(__file__), '..', 'output', 'flowchart_demo.pptx')
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    try:
        engine.create_presentation(config, output_path)
        print(f"✅ Flowchart presentation created successfully: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Error creating flowchart presentation: {e}")
        return False

def test_simple_flowchart():
    """Test a simple programmatic flowchart creation."""
    
    # Simple flowchart configuration
    simple_config = {
        "presentation": {
            "properties": {
                "title": "Simple Flowchart Test",
                "author": "PPTX Engine"
            },
            "slides": [
                {
                    "layout": 6,
                    "background": "#ffffff",
                    "shapes": [
                        {
                            "type": "text",
                            "text": "Simple Flowchart Test",
                            "x": 1,
                            "y": 0.5,
                            "w": 14,
                            "h": 1,
                            "font": {"size": 32, "bold": True, "color": "#2c3e50"},
                            "paragraph": {"alignment": "center"}
                        },
                        {
                            "type": "flowchart",
                            "x": 0,
                            "y": 0,
                            "w": 16,
                            "h": 9,
                            "elements": [
                                {
                                    "id": "start",
                                    "flowchart_type": "start",
                                    "text": "Start",
                                    "x": 7,
                                    "y": 2,
                                    "w": 2,
                                    "h": 1
                                },
                                {
                                    "id": "process1",
                                    "flowchart_type": "process",
                                    "text": "Process Step",
                                    "x": 6.5,
                                    "y": 4,
                                    "w": 3,
                                    "h": 1
                                },
                                {
                                    "id": "end",
                                    "flowchart_type": "end",
                                    "text": "End",
                                    "x": 7,
                                    "y": 6,
                                    "w": 2,
                                    "h": 1
                                }
                            ],
                            "connections": [
                                {
                                    "from": "start",
                                    "to": "process1",
                                    "connector_type": "STRAIGHT"
                                },
                                {
                                    "from": "process1",
                                    "to": "end",
                                    "connector_type": "STRAIGHT"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    }
    
    # Create the engine
    engine = PPTXEngine()
    
    # Generate the presentation
    output_path = os.path.join(os.path.dirname(__file__), '..', 'output', 'simple_flowchart.pptx')
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    try:
        engine.create_presentation(simple_config, output_path)
        print(f"✅ Simple flowchart created successfully: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Error creating simple flowchart: {e}")
        return False

if __name__ == "__main__":
    print("🔄 Testing flowchart functionality...")
    
    # Test simple flowchart first
    print("\n1. Testing simple flowchart...")
    simple_success = test_simple_flowchart()
    
    # Test comprehensive flowchart example
    print("\n2. Testing comprehensive flowchart example...")
    comprehensive_success = test_flowchart()
    
    if simple_success and comprehensive_success:
        print("\n🎉 All flowchart tests passed!")
    else:
        print("\n⚠️  Some tests failed. Check the error messages above.")
