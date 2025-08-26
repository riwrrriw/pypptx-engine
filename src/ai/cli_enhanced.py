#!/usr/bin/env python3
"""
Enhanced CLI for AI/RAG system
Generates professional presentations with LLM-powered content and image integration
"""

import argparse
import sys
import os
import json
import subprocess
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

# Add parent directories to path for imports
sys.path.append(str(Path(__file__).parent.parent))
sys.path.append(str(Path(__file__).parent.parent / "pypptx_engine"))

from ai.content_extractor import ContentExtractor
from json_generator_enhanced import EnhancedJSONGenerator
from enhanced_content_generator import EnhancedContentGenerator


def create_result_directory(work_name: str) -> Path:
    """Create result directory for the project"""
    result_dir = Path(__file__).parent.parent / "result" / work_name
    result_dir.mkdir(parents=True, exist_ok=True)
    return result_dir


def step1_generate_content_plan(work_name: str, resource_dir: Path) -> Path:
    """Step 1: Generate enhanced content plan from resources using LLM"""
    print("🔍 Step 1: Analyzing resources and generating content plan...")
    
    # Extract content from resources
    extractor = ContentExtractor(resource_dir)
    # Get extracted content by parsing the content plan
    content_plan_basic = extractor.extract_and_plan()
    
    # Convert to dictionary format for enhanced generator
    extracted_content = {"content_plan": content_plan_basic}
    
    if not extracted_content:
        raise ValueError(f"No supported files found in {resource_dir}")
    
    # Generate enhanced content plan using LLM
    result_dir = create_result_directory(work_name)
    content_plan_path = result_dir / "content_plan.md"
    
    # Use enhanced content generator with LLM
    enhanced_generator = EnhancedContentGenerator()
    content_plan = enhanced_generator.generate_enhanced_content_plan(extracted_content, work_name)
    
    with open(content_plan_path, 'w', encoding='utf-8') as f:
        f.write(content_plan)
    
    print(f"✅ Content plan generated: {content_plan_path}")
    print("📝 Please review and customize the content plan before proceeding to step 2")
    
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
    """Step 3: Generate PowerPoint presentation from JSON specification"""
    print("🎨 Step 3: Generating PowerPoint presentation...")
    
    if not json_path.exists():
        raise FileNotFoundError(f"JSON specification not found: {json_path}")
    
    # Generate PPTX using pypptx-engine
    result_dir = json_path.parent
    pptx_path = result_dir / f"{work_name}.pptx"
    
    try:
        # Use subprocess to call pypptx-engine CLI
        cmd = [
            sys.executable, "-m", "pypptx_engine.cli",
            "--input", str(json_path),
            "--output", str(pptx_path)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        
        print(f"✅ PowerPoint presentation generated: {pptx_path}")
        
        # Generate enhanced project summary
        summary_path = generate_enhanced_project_summary(work_name, result_dir, json_path)
        print(f"📋 Project summary generated: {summary_path}")
        
        return pptx_path
        
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"❌ Error generating PPTX: {e}")
    except Exception as e:
        raise RuntimeError(f"❌ Error: {e}")


def generate_enhanced_project_summary(work_name: str, result_dir: Path, json_path: Path) -> Path:
    """Generate enhanced project summary using LLM"""
    summary_path = result_dir / "project_summary.md"
    content_plan_path = result_dir / "content_plan.md"
    
    # Load JSON to analyze content
    with open(json_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    
    # Use enhanced content generator for better summary
    try:
        enhanced_generator = EnhancedContentGenerator()
        summary_content = enhanced_generator.generate_enhanced_project_summary(
            content_plan_path, json_data, work_name
        )
    except Exception as e:
        print(f"LLM summary generation failed, using fallback: {e}")
        # Fallback to basic summary
        slides_count = len(json_data.get("presentation", {}).get("slides", []))
        summary_content = f"""# Project Summary: {work_name}

## Generated Files
- **Content Plan**: content_plan.md
- **JSON Specification**: {work_name}_presentation.json
- **PowerPoint Presentation**: {work_name}.pptx
- **Project Summary**: project_summary.md

## Presentation Overview
- **Total Slides**: {slides_count}
- **Generated Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Next Steps
1. Review the generated presentation: `{work_name}.pptx`
2. Make any necessary adjustments to the content_plan.md
3. Re-run steps 2-3 if changes are needed
4. Use the presentation for your intended purpose

---
*Generated by pypptx-engine AI system with enhanced LLM content generation*
"""
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write(summary_content)
    
    return summary_path


def main():
    """Main CLI function"""
    parser = argparse.ArgumentParser(
        description="Enhanced AI/RAG system for generating professional presentations"
    )
    
    parser.add_argument("work_name", help="Name of the work/project")
    parser.add_argument("--step", type=int, choices=[1, 2, 3], help="Run specific step")
    parser.add_argument("--all", action="store_true", help="Run all steps")
    parser.add_argument("--resource-dir", type=Path, help="Directory containing resource files")
    parser.add_argument("--content-plan", type=Path, help="Path to existing content plan")
    parser.add_argument("--json-spec", type=Path, help="Path to existing JSON specification")
    
    args = parser.parse_args()
    
    try:
        if args.all:
            # Run all steps
            resource_dir = args.resource_dir or Path(__file__).parent / "resource"
            
            # Step 1
            content_plan_path = step1_generate_content_plan(args.work_name, resource_dir)
            
            # Step 2
            json_path = step2_generate_json(args.work_name, content_plan_path)
            
            # Step 3
            pptx_path = step3_generate_pptx(args.work_name, json_path)
            
            print(f"\n🎉 Complete! Generated presentation: {pptx_path}")
            
        elif args.step == 1:
            resource_dir = args.resource_dir or Path(__file__).parent / "resource"
            step1_generate_content_plan(args.work_name, resource_dir)
            
        elif args.step == 2:
            if args.content_plan:
                content_plan_path = args.content_plan
            else:
                result_dir = Path(__file__).parent.parent / "result" / args.work_name
                content_plan_path = result_dir / "content_plan.md"
            
            step2_generate_json(args.work_name, content_plan_path)
            
        elif args.step == 3:
            if args.json_spec:
                json_path = args.json_spec
            else:
                result_dir = Path(__file__).parent.parent / "result" / args.work_name
                json_path = result_dir / f"{args.work_name}_presentation.json"
            
            step3_generate_pptx(args.work_name, json_path)
            
        else:
            parser.print_help()
            
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
