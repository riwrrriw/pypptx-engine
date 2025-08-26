"""
Debug test script to identify presentation generation issues
"""
import json
import traceback
from pathlib import Path
from .engine import PPTXEngine

def test_comprehensive_example():
    """Test comprehensive example with detailed error reporting."""
    
    try:
        # Load the JSON
        json_path = Path("src/examples/comprehensive_example.json")
        with open(json_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print("✅ JSON loaded successfully")
        print(f"📊 Found {len(config['presentation']['slides'])} slides")
        
        # Initialize engine
        engine = PPTXEngine()
        print("✅ Engine initialized")
        
        # Test slide by slide
        for i, slide_config in enumerate(config['presentation']['slides']):
            print(f"\n🔍 Testing Slide {i+1}...")
            
            # Check slide structure
            shapes = slide_config.get('shapes', [])
            print(f"   - {len(shapes)} shapes")
            
            # Check each shape
            for j, shape in enumerate(shapes):
                shape_type = shape.get('type', 'unknown')
                print(f"   - Shape {j+1}: {shape_type}")
                
                # Validate shape coordinates
                coords = ['x', 'y', 'w', 'h']
                missing = [c for c in coords if c not in shape]
                if missing:
                    print(f"     ⚠️  Missing coordinates: {missing}")
                
                # Check specific shape issues
                if shape_type == 'chart':
                    chart_type = shape.get('chartType', 'unknown')
                    series_count = len(shape.get('series', []))
                    print(f"     - Chart type: {chart_type}, Series: {series_count}")
                    
                elif shape_type == 'image':
                    image_path = shape.get('path', 'unknown')
                    print(f"     - Image path: {image_path}")
                    
                elif shape_type == 'autoshape':
                    shape_subtype = shape.get('shape_type', 'unknown')
                    print(f"     - AutoShape type: {shape_subtype}")
                    
                elif shape_type == 'connector':
                    connector_type = shape.get('connector_type', 'unknown')
                    print(f"     - Connector type: {connector_type}")
        
        # Try to generate presentation
        print(f"\n🚀 Attempting to generate presentation...")
        output_path = "src/output/debug_detailed.pptx"
        engine.generate_presentation(config, output_path, "src/examples")
        print(f"✅ Presentation generated successfully: {output_path}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error during generation: {e}")
        print(f"📋 Full traceback:")
        traceback.print_exc()
        return False

def test_individual_features():
    """Test individual features separately."""
    
    # Test basic slide
    basic_config = {
        "presentation": {
            "properties": {"title": "Test", "author": "Debug"},
            "slides": [{
                "layout": 6,
                "shapes": [{
                    "type": "text",
                    "text": "Test Text",
                    "x": 1, "y": 1, "w": 4, "h": 1
                }]
            }]
        }
    }
    
    try:
        engine = PPTXEngine()
        engine.generate_presentation(basic_config, "src/output/basic_test.pptx", "src/examples")
        print("✅ Basic slide test passed")
        return True
    except Exception as e:
        print(f"❌ Basic slide test failed: {e}")
        return False

if __name__ == "__main__":
    print("🔧 pypptx-engine Debug Test\n")
    
    # Test basic functionality first
    print("1️⃣ Testing basic functionality...")
    basic_ok = test_individual_features()
    
    if basic_ok:
        print("\n2️⃣ Testing comprehensive example...")
        comprehensive_ok = test_comprehensive_example()
        
        if comprehensive_ok:
            print("\n🎉 All tests passed!")
        else:
            print("\n⚠️ Comprehensive example has issues")
    else:
        print("\n❌ Basic functionality failed")
