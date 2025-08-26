"""
Image Service for AI/RAG system
Integrates with Unsplash and Pexels APIs for professional image selection
"""

import requests
import json
from typing import List, Dict, Any, Optional
from pathlib import Path
import os


class ImageService:
    """Professional image service with Unsplash and Pexels integration"""
    
    def __init__(self):
        # API Keys from environment
        self.unsplash_access_key = "3rV7jP1d2jUZJhSaKq7aC9kTRQsWtRv-iGvAUzBEq_w"
        self.pexels_api_key = "93lXRhhE6ytudJB20d67RRDoUvrzNXbT42eZjXoOoIAsPEO75feEnAJW"
        
        self.unsplash_base_url = "https://api.unsplash.com"
        self.pexels_base_url = "https://api.pexels.com/v1"
    
    def search_unsplash_images(self, query: str, count: int = 5) -> List[Dict[str, Any]]:
        """Search for professional images on Unsplash"""
        try:
            headers = {
                "Authorization": f"Client-ID {self.unsplash_access_key}"
            }
            
            params = {
                "query": query,
                "per_page": count,
                "orientation": "landscape",
                "content_filter": "high"
            }
            
            response = requests.get(
                f"{self.unsplash_base_url}/search/photos",
                headers=headers,
                params=params,
                timeout=10
            )
            
            if response.status_code == 200:
                data = response.json()
                images = []
                
                for photo in data.get("results", []):
                    images.append({
                        "url": photo["urls"]["regular"],
                        "thumb_url": photo["urls"]["thumb"],
                        "description": photo.get("alt_description", query),
                        "photographer": photo["user"]["name"],
                        "source": "unsplash",
                        "width": photo["width"],
                        "height": photo["height"]
                    })
                
                return images
            
        except Exception as e:
            print(f"Unsplash API error: {e}")
        
        return []
    
    def search_pexels_images(self, query: str, count: int = 5) -> List[Dict[str, Any]]:
        """Search for professional images on Pexels"""
        try:
            headers = {
                "Authorization": self.pexels_api_key
            }
            
            params = {
                "query": query,
                "per_page": count,
                "orientation": "landscape"
            }
            
            response = requests.get(
                f"{self.pexels_base_url}/search",
                headers=headers,
                params=params,
                timeout=10
            )
            
            if response.status_code == 200:
                data = response.json()
                images = []
                
                for photo in data.get("photos", []):
                    images.append({
                        "url": photo["src"]["large"],
                        "thumb_url": photo["src"]["medium"],
                        "description": photo.get("alt", query),
                        "photographer": photo["photographer"],
                        "source": "pexels",
                        "width": photo["width"],
                        "height": photo["height"]
                    })
                
                return images
            
        except Exception as e:
            print(f"Pexels API error: {e}")
        
        return []
    
    def find_best_image(self, query: str, slide_context: str = "") -> Optional[Dict[str, Any]]:
        """Find the best professional image for a slide"""
        # Try Unsplash first (higher quality)
        images = self.search_unsplash_images(query, 3)
        
        # Fallback to Pexels if Unsplash fails
        if not images:
            images = self.search_pexels_images(query, 3)
        
        if images:
            # Return the first (best) result
            return images[0]
        
        # Fallback to a default professional image
        return {
            "url": "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&q=80",
            "description": "Professional background",
            "photographer": "Unsplash",
            "source": "default"
        }
    
    def get_slide_image_suggestions(self, slide_title: str, content_type: str) -> List[str]:
        """Get image search queries based on slide content"""
        suggestions = []
        
        # Content-based suggestions
        if "business" in slide_title.lower() or "corporate" in slide_title.lower():
            suggestions.extend(["business meeting", "corporate office", "professional team"])
        elif "technology" in slide_title.lower() or "tech" in slide_title.lower():
            suggestions.extend(["modern technology", "digital innovation", "tech workspace"])
        elif "data" in slide_title.lower() or "analytics" in slide_title.lower():
            suggestions.extend(["data visualization", "analytics dashboard", "business charts"])
        elif "growth" in slide_title.lower() or "success" in slide_title.lower():
            suggestions.extend(["business growth", "success concept", "upward trend"])
        elif "team" in slide_title.lower() or "collaboration" in slide_title.lower():
            suggestions.extend(["team collaboration", "business meeting", "teamwork"])
        
        # Content type suggestions
        if content_type == "title":
            suggestions.extend(["professional background", "modern office", "business concept"])
        elif content_type == "bullet_list":
            suggestions.extend(["checklist", "business planning", "strategy"])
        elif content_type == "comparison":
            suggestions.extend(["comparison concept", "choice decision", "options"])
        
        # Default fallbacks
        if not suggestions:
            suggestions = ["professional background", "business concept", "modern office"]
        
        return suggestions[:3]  # Return top 3 suggestions
