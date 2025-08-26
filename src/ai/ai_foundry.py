"""
AI Foundry Integration for pypptx-engine
Azure OpenAI integration using the official Python SDK
"""

import os
import json
from typing import Dict, List, Any, Optional
from pathlib import Path

try:
    from openai import AzureOpenAI
    AZURE_OPENAI_AVAILABLE = True
except ImportError:
    AZURE_OPENAI_AVAILABLE = False


class AIFoundryClient:
    """Client for Azure OpenAI service"""
    
    def __init__(self, api_key: Optional[str] = None, model: str = "gpt-5"):
        self.api_key = api_key or os.getenv("AI_FOUNDRY_API_KEY")
        self.model = model or os.getenv("AI_MODEL", "gpt-5")
        self.deployment = os.getenv("AI_MODEL_DEPLOYMENT", "gpt-5-for-pptx")
        self.endpoint = os.getenv("AI_FOUNDRY_BASE_URL", "https://ai-peerawatr-6647.cognitiveservices.azure.com/")
        self.api_version = os.getenv("AI_FOUNDRY_API_VERSION", "2024-12-01-preview")
        
        # Initialize Azure OpenAI client
        if AZURE_OPENAI_AVAILABLE and self.api_key:
            self.client = AzureOpenAI(
                api_version=self.api_version,
                azure_endpoint=self.endpoint,
                api_key=self.api_key,
            )
        else:
            self.client = None
        
        # Model configurations available through Azure OpenAI
        self.models = {
            "gpt-35-turbo": {"context_length": 4096, "type": "chat"},
            "gpt-4": {"context_length": 8192, "type": "chat"},
            "gpt-4o": {"context_length": 128000, "type": "chat"},
            "gpt-4o-mini": {"context_length": 128000, "type": "chat"},
            "gpt-5": {"context_length": 128000, "type": "chat"}
        }
    
    def enhance_content_plan(self, extracted_content: Dict[str, Any], 
                           project_context: str = "") -> str:
        """Use AI to enhance and structure content plan"""
        if not self.api_key:
            return self._fallback_content_plan(extracted_content, project_context)
        
        prompt = self._build_content_enhancement_prompt(extracted_content, project_context)
        
        try:
            response = self._call_ai_model(prompt)
            return response
        except Exception as e:
            print(f"AI enhancement failed, using fallback: {e}")
            return self._fallback_content_plan(extracted_content, project_context)
    
    def improve_slide_content(self, slide_info: Dict[str, Any], 
                            context: str = "") -> Dict[str, Any]:
        """Use AI to improve individual slide content"""
        if not self.api_key:
            return slide_info
        
        prompt = f"""
Improve this slide content for a professional presentation:

Title: {slide_info.get('title', 'Untitled')}
Current content: {slide_info.get('main_points', [])}
Context: {context}

Please provide:
1. A more engaging title
2. 3-5 clear, concise bullet points
3. Suggested visual elements
4. Professional tone and structure

Return as JSON with keys: title, main_points, visual_suggestions, notes
"""
        
        try:
            response = self._call_ai_model(prompt)
            # Parse AI response and merge with existing slide_info
            ai_suggestions = json.loads(response)
            slide_info.update(ai_suggestions)
            return slide_info
        except Exception as e:
            print(f"AI slide improvement failed: {e}")
            return slide_info
    
    def generate_design_suggestions(self, content_summary: str, 
                                  target_audience: str = "") -> Dict[str, Any]:
        """Generate design and visual suggestions using AI"""
        if not self.api_key:
            return self._default_design_suggestions()
        
        prompt = f"""
Based on this presentation content and audience, suggest design elements:

Content Summary: {content_summary}
Target Audience: {target_audience}

Provide suggestions for:
1. Color scheme (primary, secondary, accent colors)
2. Font style and sizing
3. Layout preferences
4. Image/visual style
5. Animation level
6. Overall design theme

Return as JSON with appropriate keys.
"""
        
        try:
            response = self._call_ai_model(prompt)
            return json.loads(response)
        except Exception as e:
            print(f"AI design suggestions failed: {e}")
            return self._default_design_suggestions()
    
    def _call_ai_model(self, prompt: str, max_tokens: int = 2000) -> str:
        """Call Azure OpenAI using Python SDK"""
        return self._call_azure_openai(prompt, max_tokens)
    
    def _call_azure_openai(self, prompt: str, max_tokens: int) -> str:
        """Call Azure OpenAI using the official Python SDK"""
        if not self.client:
            if not AZURE_OPENAI_AVAILABLE:
                raise ImportError("Azure OpenAI SDK not available. Install with: pip install openai")
            if not self.api_key:
                raise ValueError("Azure OpenAI API key not provided")
        
        try:
            response = self.client.chat.completions.create(
                messages=[
                    {
                        "role": "system",
                        "content": "You are a helpful assistant specialized in creating professional presentation content."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                max_completion_tokens=max_tokens,
                model=self.deployment,
                temperature=0.7
            )
            
            return response.choices[0].message.content
        except Exception as e:
            print(f"Azure OpenAI API call failed: {e}")
            raise
    
    def _build_content_enhancement_prompt(self, extracted_content: Dict[str, Any], 
                                        context: str) -> str:
        """Build prompt for content enhancement"""
        content_summary = []
        for filename, content in extracted_content.items():
            if "error" not in content:
                summary = content.get("summary", "No summary available")
                content_summary.append(f"- {filename}: {summary}")
        
        return f"""
You are an expert presentation designer. Create a comprehensive content plan for a professional PowerPoint presentation.

Available Content:
{chr(10).join(content_summary)}

Additional Context: {context}

Please create a structured content plan that includes:
1. An engaging presentation title
2. 5-8 well-organized slides with clear titles
3. Key points for each slide (3-5 bullets max)
4. Suggested visual elements and design
5. Logical flow and storytelling structure

Focus on:
- Professional, engaging content
- Clear value proposition
- Audience-appropriate language
- Visual storytelling opportunities
- Actionable insights

Return the plan in markdown format similar to a content plan template.
"""
    
    def _fallback_content_plan(self, extracted_content: Dict[str, Any], 
                             context: str) -> str:
        """Fallback content plan when AI is not available"""
        return f"""
# Enhanced Content Plan (Fallback Mode)

## AI Enhancement Note
AI Foundry integration is available but not configured. 
To enable AI enhancement, set your API key:
```bash
export AI_FOUNDRY_API_KEY="your-api-key-here"
```

## Available Content Summary
"""
    
    def _default_design_suggestions(self) -> Dict[str, Any]:
        """Default design suggestions when AI is not available"""
        return {
            "color_scheme": {
                "primary": "#1f4e79",
                "secondary": "#70ad47", 
                "accent": "#ffc000"
            },
            "font_style": "Modern and professional",
            "layout": "Clean with ample white space",
            "image_style": "Professional photography",
            "animation_level": "Subtle transitions",
            "theme": "Corporate professional"
        }


class ContentEnhancer:
    """Enhanced content processing with AI integration"""
    
    def __init__(self, ai_client: Optional[AIFoundryClient] = None):
        self.ai_client = ai_client or AIFoundryClient()
    
    def enhance_extracted_content(self, extracted_content: Dict[str, Any], 
                                project_name: str = "") -> Dict[str, Any]:
        """Enhance extracted content using AI"""
        enhanced = extracted_content.copy()
        
        for filename, content in enhanced.items():
            if "error" in content:
                continue
            
            # Enhance summary with AI
            if self.ai_client.api_key:
                try:
                    enhanced_summary = self._enhance_summary(content, filename)
                    content["ai_enhanced_summary"] = enhanced_summary
                except Exception as e:
                    print(f"Failed to enhance {filename}: {e}")
            
            # Extract key topics
            content["key_topics"] = self._extract_key_topics(content)
            
            # Suggest slide types
            content["suggested_slide_types"] = self._suggest_slide_types(content)
        
        return enhanced
    
    def _enhance_summary(self, content: Dict[str, Any], filename: str) -> str:
        """Use AI to create better summaries"""
        text = content.get("text", "")
        if not text:
            return content.get("summary", "No content available")
        
        prompt = f"""
Summarize this content for a professional presentation slide:

File: {filename}
Content: {text[:1000]}...

Provide a concise, engaging summary (2-3 sentences) that would work well as slide content.
Focus on key insights and actionable information.
"""
        
        return self.ai_client._call_ai_model(prompt, max_tokens=150)
    
    def _extract_key_topics(self, content: Dict[str, Any]) -> List[str]:
        """Extract key topics from content"""
        text = content.get("text", "")
        if not text:
            return []
        
        # Simple keyword extraction (can be enhanced with NLP)
        words = text.lower().split()
        
        # Filter common words and find important terms
        stop_words = {"the", "and", "or", "but", "in", "on", "at", "to", "for", "of", "with", "by"}
        important_words = [word for word in words if len(word) > 4 and word not in stop_words]
        
        # Return most frequent terms (simplified approach)
        from collections import Counter
        word_counts = Counter(important_words)
        return [word for word, count in word_counts.most_common(5)]
    
    def _suggest_slide_types(self, content: Dict[str, Any]) -> List[str]:
        """Suggest appropriate slide types based on content"""
        text = content.get("text", "").lower()
        suggestions = []
        
        if "step" in text or "process" in text or "how to" in text:
            suggestions.append("flowchart")
        
        if "data" in text or "statistics" in text or "%" in text:
            suggestions.append("chart")
        
        if "compare" in text or "versus" in text or "vs" in text:
            suggestions.append("table")
        
        if "image" in text or "photo" in text or "picture" in text:
            suggestions.append("image")
        
        if not suggestions:
            suggestions.append("bullet_list")
        
        return suggestions
