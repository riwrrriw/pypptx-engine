#!/usr/bin/env python3
"""
AI-powered CLI for pypptx-engine
Generates presentations through a 3-step process:
1. Extract content from resources and generate content_plan.md
2. Generate JSON specification from content plan
3. Create PPTX using pypptx-engine
"""

import os
import sys
import json
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent))
sys.path.append(str(Path(__file__).parent.parent / "pypptx_engine"))

from ai.content_extractor import ContentExtractor
from json_generator_enhanced import EnhancedJSONGenerator


def create_result_directory(work_name: str) -> Path:
    """Create result directory for the project"""
    result_dir = Path(__file__).parent.parent / "result" / work_name
    result_dir.mkdir(parents=True, exist_ok=True)
    return result_dir


def step1_generate_content_plan(work_name: str, resource_dir: Path) -> Path:
    """Step 1: Generate content_plan.md from resources"""
    print("🔍 Step 1: Analyzing resources and generating content plan...")
    
    extractor = ContentExtractor(resource_dir)
    content_plan = extractor.extract_and_plan()
    
    # Save content plan
    result_dir = create_result_directory(work_name)
    content_plan_path = result_dir / "content_plan.md"
    
    with open(content_plan_path, 'w', encoding='utf-8') as f:
        f.write(content_plan)
    
    print(f"✅ Content plan generated: {content_plan_path}")
    print(f"📝 Please review and customize the content plan before proceeding to step 2")
    
    return content_plan_path


def step2_generate_json(work_name: str, content_plan_path: Path) -> Path:
    """Step 2: Generate JSON specification from content plan"""
    print("🔧 Step 2: Generating JSON specification from content plan...")
    
    if not content_plan_path.exists():
        raise FileNotFoundError(f"Content plan not found: {content_plan_path}")
    
    generator = EnhancedJSONGenerator()
    json_spec = generator.generate_from_content_plan(content_plan_path)
    
    # Save JSON specification
    result_dir = Path(content_plan_path).parent
    json_path = result_dir / f"{work_name}_presentation.json"
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_spec, f, indent=2, ensure_ascii=False)
    
    print(f"✅ JSON specification generated: {json_path}")
    
    return json_path


def step3_generate_pptx(work_name: str, json_path: Path) -> Path:
    """Step 3: Generate PPTX from JSON using pypptx-engine"""
    print("🎨 Step 3: Generating PowerPoint presentation...")
    
    if not json_path.exists():
        raise FileNotFoundError(f"JSON specification not found: {json_path}")
    
    result_dir = Path(json_path).parent
    pptx_path = result_dir / f"{work_name}.pptx"
    
    # Use pypptx-engine directly with Python module execution
    main_project_dir = Path(__file__).parent.parent.parent
    
    cmd = [
        sys.executable, "-m", "pypptx_engine.cli",
        "--input", str(json_path),
        "--output", str(pptx_path)
    ]
    
    try:
        # Set PYTHONPATH to include src directory
        env = os.environ.copy()
        env["PYTHONPATH"] = str(main_project_dir / "src")
        
        subprocess.run(cmd, check=True, cwd=str(main_project_dir), env=env)
        print(f"✅ PowerPoint presentation generated: {pptx_path}")
        
        # Generate summary
        generate_project_summary(work_name, result_dir, json_path)
        
        return pptx_path
    except subprocess.CalledProcessError as e:
        print(f"❌ Error generating PPTX: {e}")
        raise


def generate_project_summary(work_name: str, result_dir: Path, json_path: Path):
    """Generate a summary of the project"""
    summary_path = result_dir / "project_summary.md"
    
    # Load JSON to analyze content
    with open(json_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    slides_info = []
    for i, slide in enumerate(json_data.get('slides', []), 1):
        slide_info = {
            'number': i,
            'title': slide.get('title', f'Slide {i}'),
            'content_types': [],
            'resources': []
        }
        
        for element in slide.get('elements', []):
            element_type = element.get('type', 'unknown')
            if element_type not in slide_info['content_types']:
                slide_info['content_types'].append(element_type)
        
        slides_info.append(slide_info)
    
    summary_content = f"""# Project Summary: {work_name}

## Generated Files
- **Content Plan**: content_plan.md
- **JSON Specification**: {json_path.name}
- **PowerPoint Presentation**: {work_name}.pptx
- **Project Summary**: project_summary.md

## Presentation Overview
- **Total Slides**: {len(slides_info)}
- **Generated Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Slide Content Overview
"""
    
    for slide in slides_info:
        summary_content += f"""
### Slide {slide['number']}: {slide['title']}
- **Content Types**: {', '.join(slide['content_types']) if slide['content_types'] else 'None'}
"""
    
    summary_content += f"""
## Resource Mapping
*Check content_plan.md for detailed resource mapping*

## Next Steps
1. Review the generated presentation: `{work_name}.pptx`
2. Make any necessary adjustments to the content_plan.md
3. Re-run steps 2-3 if changes are needed
4. Use the presentation for your intended purpose

---
*Generated by pypptx-engine AI system*
"""
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write(summary_content)
    
    print(f"📋 Project summary generated: {summary_path}")


def main():
    parser = argparse.ArgumentParser(
        description="AI-powered presentation generator for pypptx-engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full 3-step process
  python cli.py my_project --all
  
  # Step by step
  python cli.py my_project --step 1
  python cli.py my_project --step 2
  python cli.py my_project --step 3
  
  # Skip to specific step
  python cli.py my_project --step 2 --content-plan /path/to/content_plan.md
        """
    )
    
    parser.add_argument("work_name", help="Name of the work/project")
    parser.add_argument("--step", type=int, choices=[1, 2, 3], 
                       help="Run specific step (1: content plan, 2: JSON, 3: PPTX)")
    parser.add_argument("--all", action="store_true", 
                       help="Run all steps sequentially")
    parser.add_argument("--resource-dir", type=Path,
                       help="Path to resource directory (default: ./resource/)")
    parser.add_argument("--content-plan", type=Path,
                       help="Path to existing content plan (for step 2)")
    parser.add_argument("--json-spec", type=Path,
                       help="Path to existing JSON specification (for step 3)")
    
    args = parser.parse_args()
    
    if not args.step and not args.all:
        parser.error("Must specify either --step or --all")
    
    # Set default resource directory
    if not args.resource_dir:
        args.resource_dir = Path(__file__).parent / "resource"
    
    try:
        if args.all:
            print(f"🚀 Starting full AI presentation generation for: {args.work_name}")
            print("=" * 60)
            
            # Step 1
            content_plan_path = step1_generate_content_plan(args.work_name, args.resource_dir)
            
            print("\n" + "=" * 60)
            input("Press Enter to continue to step 2 (after reviewing content plan)...")
            
            # Step 2
            json_path = step2_generate_json(args.work_name, content_plan_path)
            
            print("\n" + "=" * 60)
            
            # Step 3
            pptx_path = step3_generate_pptx(args.work_name, json_path)
            
            print("\n" + "🎉" * 20)
            print(f"✅ Complete! Your presentation is ready: {pptx_path}")
            
        elif args.step == 1:
            step1_generate_content_plan(args.work_name, args.resource_dir)
            
        elif args.step == 2:
            if args.content_plan:
                content_plan_path = args.content_plan
            else:
                result_dir = create_result_directory(args.work_name)
                content_plan_path = result_dir / "content_plan.md"
            
            step2_generate_json(args.work_name, content_plan_path)
            
        elif args.step == 3:
            if args.json_spec:
                json_path = args.json_spec
            else:
                result_dir = create_result_directory(args.work_name)
                json_path = result_dir / f"{args.work_name}_presentation.json"
            
            step3_generate_pptx(args.work_name, json_path)
    
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
