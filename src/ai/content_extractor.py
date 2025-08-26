"""
Content Extractor for AI/RAG system
Extracts content from various document formats in the resource directory
"""

import os
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
import mimetypes

try:
    import PyPDF2
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import markdown
    MARKDOWN_AVAILABLE = True
except ImportError:
    MARKDOWN_AVAILABLE = False


class ContentExtractor:
    """Extract and analyze content from resource files"""
    
    def __init__(self, resource_dir: Path):
        self.resource_dir = Path(resource_dir)
        self.supported_formats = {
            '.pdf': self._extract_pdf,
            '.docx': self._extract_docx,
            '.doc': self._extract_docx,
            '.txt': self._extract_text,
            '.md': self._extract_markdown,
            '.json': self._extract_json
        }
    
    def extract_and_plan(self) -> str:
        """Extract content from all resources and generate content plan"""
        if not self.resource_dir.exists():
            raise FileNotFoundError(f"Resource directory not found: {self.resource_dir}")
        
        # Find all resource files
        resource_files = self._find_resource_files()
        
        if not resource_files:
            return self._generate_empty_plan()
        
        # Extract content from each file
        extracted_content = {}
        for file_path in resource_files:
            try:
                content = self._extract_file_content(file_path)
                extracted_content[file_path.name] = content
            except Exception as e:
                print(f"Warning: Could not extract from {file_path.name}: {e}")
                extracted_content[file_path.name] = {"error": str(e)}
        
        # Generate content plan
        return self._generate_content_plan(extracted_content, resource_files)
    
    def _find_resource_files(self) -> List[Path]:
        """Find all supported resource files"""
        files = []
        for file_path in self.resource_dir.rglob('*'):
            if file_path.is_file() and file_path.suffix.lower() in self.supported_formats:
                files.append(file_path)
        return sorted(files)
    
    def _extract_file_content(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from a single file"""
        suffix = file_path.suffix.lower()
        if suffix in self.supported_formats:
            return self.supported_formats[suffix](file_path)
        else:
            return {"error": f"Unsupported format: {suffix}"}
    
    def _extract_pdf(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from PDF file"""
        if not PDF_AVAILABLE:
            return {"error": "PyPDF2 not installed. Run: pip install PyPDF2"}
        
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"
                
                return {
                    "type": "pdf",
                    "pages": len(reader.pages),
                    "text": text.strip(),
                    "summary": self._summarize_text(text)
                }
        except Exception as e:
            return {"error": f"PDF extraction failed: {e}"}
    
    def _extract_docx(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from Word document"""
        if not DOCX_AVAILABLE:
            return {"error": "python-docx not installed. Run: pip install python-docx"}
        
        try:
            doc = Document(file_path)
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            text = "\n".join(paragraphs)
            
            return {
                "type": "docx",
                "paragraphs": len(paragraphs),
                "text": text,
                "summary": self._summarize_text(text)
            }
        except Exception as e:
            return {"error": f"DOCX extraction failed: {e}"}
    
    def _extract_text(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from text file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
                
                return {
                    "type": "text",
                    "lines": len(text.splitlines()),
                    "text": text,
                    "summary": self._summarize_text(text)
                }
        except Exception as e:
            return {"error": f"Text extraction failed: {e}"}
    
    def _extract_markdown(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from Markdown file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
                
                # Extract headers for structure
                headers = []
                for line in text.splitlines():
                    if line.startswith('#'):
                        headers.append(line.strip())
                
                return {
                    "type": "markdown",
                    "headers": headers,
                    "text": text,
                    "summary": self._summarize_text(text)
                }
        except Exception as e:
            return {"error": f"Markdown extraction failed: {e}"}
    
    def _extract_json(self, file_path: Path) -> Dict[str, Any]:
        """Extract content from JSON file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                
                return {
                    "type": "json",
                    "structure": self._analyze_json_structure(data),
                    "data": data,
                    "summary": f"JSON file with {len(data) if isinstance(data, (list, dict)) else 1} items"
                }
        except Exception as e:
            return {"error": f"JSON extraction failed: {e}"}
    
    def _analyze_json_structure(self, data: Any) -> str:
        """Analyze JSON structure"""
        if isinstance(data, dict):
            return f"Object with keys: {list(data.keys())}"
        elif isinstance(data, list):
            return f"Array with {len(data)} items"
        else:
            return f"Value of type {type(data).__name__}"
    
    def _summarize_text(self, text: str) -> str:
        """Create a simple summary of text content"""
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        if not lines:
            return "Empty content"
        
        # Get first few meaningful lines as summary
        summary_lines = []
        for line in lines[:5]:
            if len(line) > 10:  # Skip very short lines
                summary_lines.append(line)
                if len(summary_lines) >= 3:
                    break
        
        summary = " ".join(summary_lines)
        if len(summary) > 200:
            summary = summary[:200] + "..."
        
        return summary or "Content available but no clear summary"
    
    def _generate_content_plan(self, extracted_content: Dict[str, Any], resource_files: List[Path]) -> str:
        """Generate content plan from extracted content"""
        project_name = self.resource_dir.parent.name or "AI Generated Presentation"
        
        # Analyze content to suggest slides
        total_content_items = sum(1 for content in extracted_content.values() if "error" not in content)
        estimated_slides = min(max(total_content_items + 1, 3), 12)  # 3-12 slides
        
        plan = f"""# Content Plan Template

## Project Information
- **Project Name**: {project_name}
- **Generated Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- **Source Resources**: {', '.join([f.name for f in resource_files])}

## Presentation Overview
- **Title**: {project_name}
- **Target Audience**: [Please specify your target audience]
- **Presentation Goal**: [Please specify the main goal]
- **Estimated Duration**: [Please specify duration]
- **Total Slides**: {estimated_slides}

## Content Structure

### Slide 1: Title Slide
- **Title**: {project_name}
- **Subtitle**: [Auto-generated from resource analysis]
- **Background**: gradient
- **Source**: Generated

"""
        
        # Generate slide suggestions based on content
        slide_num = 2
        for filename, content in extracted_content.items():
            if "error" in content:
                continue
            
            summary = content.get('summary', 'Content from ' + filename)
            
            plan += f"""### Slide {slide_num}: {filename.replace('_', ' ').replace('.', ' ').title()}
- **Content Type**: text
- **Main Points**: 
  - {summary}
  - [Add more points as needed]
- **Visual Elements**: [Specify images, charts, or diagrams needed]
- **Background**: [Choose styling preferences]
- **Source**: {filename}

"""
            slide_num += 1
            
            if slide_num > estimated_slides:
                break
        
        plan += f"""## Design Preferences
- **Color Scheme**: Professional blue and white
- **Font Style**: Modern and clean
- **Animation Level**: Subtle
- **Image Style**: Professional photography

## Resource Mapping
| Slide | Resource File | Content Extracted | Summary |
|-------|---------------|-------------------|---------|
"""
        
        # Add resource mapping table
        slide_num = 2
        for filename, content in extracted_content.items():
            if "error" in content:
                plan += f"| N/A   | {filename}    | Error: {content['error']} | Failed to process |\n"
            else:
                summary = content.get('summary', 'Processed successfully')
                plan += f"| {slide_num}     | {filename}    | {content.get('type', 'unknown')} content | {summary[:50]}... |\n"
                slide_num += 1
        
        plan += f"""
## Notes for JSON Generation
- Focus on clean, professional design
- Use consistent formatting across slides
- Include appropriate transitions between slides
- Ensure readability with good contrast
- Consider the target audience when selecting content

---
*This content plan can be edited before proceeding to JSON generation*
*Review and customize the content above, then run step 2 to generate JSON*
"""
        
        return plan
    
    def _generate_empty_plan(self) -> str:
        """Generate a basic plan when no resources are found"""
        return f"""# Content Plan Template

## Project Information
- **Project Name**: New Presentation
- **Generated Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- **Source Resources**: No resources found in {self.resource_dir}

## Instructions
1. Add your resource files (PDF, DOCX, TXT, MD, JSON) to the resource/ directory
2. Run step 1 again to extract content and generate a proper content plan
3. Supported formats: .pdf, .docx, .doc, .txt, .md, .json

## Required Dependencies
For full functionality, install:
```bash
pip install PyPDF2 python-docx markdown
```

---
*Add resources and re-run to generate a complete content plan*
"""
