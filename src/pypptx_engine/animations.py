"""
Animation and transition system for PPTX Engine.
Provides slide transitions and shape animations.
"""

from typing import Dict, Any, List, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
import warnings


class AnimationManager:
    """Manages slide transitions and shape animations."""
    
    def __init__(self):
        self.transition_types = {
            "none": 0,
            "fade": 1,
            "push": 2,
            "wipe": 3,
            "split": 4,
            "reveal": 5,
            "random_bars": 6,
            "shape": 7,
            "uncover": 8,
            "cover": 9,
            "cut": 10,
            "fade_through_black": 11,
            "zoom": 12,
            "fly_through": 13,
            "rotate": 14,
            "newsflash": 15,
            "alpha": 16,
            "cube": 17,
            "flip": 18,
            "gallery": 19,
            "conveyor": 20,
            "pan": 21,
            "glitter": 22,
            "honeycomb": 23,
            "flash": 24,
            "shred": 25
        }
        
        self.animation_types = {
            # Entrance animations
            "appear": {"category": "entrance", "effect": 1},
            "fade_in": {"category": "entrance", "effect": 2},
            "fly_in": {"category": "entrance", "effect": 3},
            "float_in": {"category": "entrance", "effect": 4},
            "split": {"category": "entrance", "effect": 5},
            "wipe": {"category": "entrance", "effect": 6},
            "shape": {"category": "entrance", "effect": 7},
            "wheel": {"category": "entrance", "effect": 8},
            "random_bars": {"category": "entrance", "effect": 9},
            "grow_and_turn": {"category": "entrance", "effect": 10},
            "zoom": {"category": "entrance", "effect": 11},
            "swivel": {"category": "entrance", "effect": 12},
            "bounce": {"category": "entrance", "effect": 13},
            
            # Emphasis animations
            "pulse": {"category": "emphasis", "effect": 1},
            "color_pulse": {"category": "emphasis", "effect": 2},
            "teeter": {"category": "emphasis", "effect": 3},
            "spin": {"category": "emphasis", "effect": 4},
            "grow_shrink": {"category": "emphasis", "effect": 5},
            "desaturate": {"category": "emphasis", "effect": 6},
            "darken": {"category": "emphasis", "effect": 7},
            "lighten": {"category": "emphasis", "effect": 8},
            "transparency": {"category": "emphasis", "effect": 9},
            "object_color": {"category": "emphasis", "effect": 10},
            "complementary_color": {"category": "emphasis", "effect": 11},
            "line_color": {"category": "emphasis", "effect": 12},
            "fill_color": {"category": "emphasis", "effect": 13},
            
            # Exit animations
            "disappear": {"category": "exit", "effect": 1},
            "fade_out": {"category": "exit", "effect": 2},
            "fly_out": {"category": "exit", "effect": 3},
            "float_out": {"category": "exit", "effect": 4},
            "split_out": {"category": "exit", "effect": 5},
            "wipe_out": {"category": "exit", "effect": 6},
            "shape_out": {"category": "exit", "effect": 7},
            "random_bars_out": {"category": "exit", "effect": 8},
            "shrink_and_turn": {"category": "exit", "effect": 9},
            "zoom_out": {"category": "exit", "effect": 10},
            "swivel_out": {"category": "exit", "effect": 11},
            "bounce_out": {"category": "exit", "effect": 12}
        }
        
        self.timing_options = {
            "very_fast": 0.5,
            "fast": 1.0,
            "medium": 2.0,
            "slow": 3.0,
            "very_slow": 5.0
        }
    
    def apply_slide_transition(self, slide, transition_config: Dict[str, Any]) -> None:
        """Apply transition effect to slide."""
        try:
            transition_type = transition_config.get("type", "none")
            duration = transition_config.get("duration", 1.0)
            
            # Convert string duration to float if needed
            if isinstance(duration, str):
                duration = self.timing_options.get(duration, 1.0)
            
            # Apply transition using python-pptx (limited support)
            # Note: python-pptx has limited transition support, this is a basic implementation
            if hasattr(slide, 'slide_layout') and hasattr(slide.slide_layout, 'slide_master'):
                # Store transition info in slide notes for reference
                if not hasattr(slide, 'notes_slide') or slide.notes_slide is None:
                    slide.notes_slide = slide.part.package.presentation_part.presentation.slides._sldIdLst[-1].get_or_add_notes_slide()
                
                # Add transition metadata to notes
                transition_info = f"Transition: {transition_type}, Duration: {duration}s"
                if slide.notes_slide.notes_text_frame.text:
                    slide.notes_slide.notes_text_frame.text += f"\n{transition_info}"
                else:
                    slide.notes_slide.notes_text_frame.text = transition_info
                    
        except Exception as e:
            warnings.warn(f"Failed to apply slide transition: {e}")
    
    def apply_shape_animation(self, shape, animation_config: Dict[str, Any]) -> None:
        """Apply animation to shape."""
        try:
            animation_type = animation_config.get("type", "appear")
            delay = animation_config.get("delay", 0)
            duration = animation_config.get("duration", 1.0)
            trigger = animation_config.get("trigger", "on_click")
            
            # Convert string duration to float if needed
            if isinstance(duration, str):
                duration = self.timing_options.get(duration, 1.0)
            
            # Store animation metadata in shape name for reference
            # Note: python-pptx has very limited animation support
            animation_info = f"anim:{animation_type}:{duration}:{delay}:{trigger}"
            
            if hasattr(shape, 'name'):
                if shape.name:
                    shape.name = f"{shape.name}|{animation_info}"
                else:
                    shape.name = animation_info
            
            # Apply visual effects that can be simulated
            self._apply_visual_effects(shape, animation_config)
                    
        except Exception as e:
            warnings.warn(f"Failed to apply shape animation: {e}")
    
    def _apply_visual_effects(self, shape, animation_config: Dict[str, Any]) -> None:
        """Apply visual effects that can be simulated with formatting."""
        animation_type = animation_config.get("type", "appear")
        
        # Apply initial state based on animation type
        if animation_type in ["fade_in", "appear"]:
            # For fade in, we could start with low transparency
            # but python-pptx transparency support is limited
            pass
        elif animation_type == "zoom":
            # For zoom, we could adjust initial size
            # This would require storing original dimensions
            pass
        elif animation_type in ["fly_in", "float_in"]:
            # For fly in, we could adjust initial position
            # This would require storing target position
            pass
    
    def create_animation_sequence(self, shapes: List[Any], sequence_config: Dict[str, Any]) -> None:
        """Create animation sequence for multiple shapes."""
        sequence_type = sequence_config.get("type", "sequential")
        base_delay = sequence_config.get("base_delay", 0)
        delay_increment = sequence_config.get("delay_increment", 0.5)
        
        for i, shape in enumerate(shapes):
            if sequence_type == "sequential":
                delay = base_delay + (i * delay_increment)
            elif sequence_type == "simultaneous":
                delay = base_delay
            elif sequence_type == "reverse":
                delay = base_delay + ((len(shapes) - 1 - i) * delay_increment)
            else:
                delay = base_delay
            
            # Apply animation with calculated delay
            animation_config = sequence_config.get("animation", {}).copy()
            animation_config["delay"] = delay
            
            self.apply_shape_animation(shape, animation_config)
    
    def get_available_transitions(self) -> List[str]:
        """Get list of available transition types."""
        return list(self.transition_types.keys())
    
    def get_available_animations(self) -> Dict[str, List[str]]:
        """Get list of available animations by category."""
        animations_by_category = {
            "entrance": [],
            "emphasis": [],
            "exit": []
        }
        
        for anim_name, anim_info in self.animation_types.items():
            category = anim_info["category"]
            animations_by_category[category].append(anim_name)
        
        return animations_by_category
    
    def validate_animation_config(self, config: Dict[str, Any]) -> bool:
        """Validate animation configuration."""
        if "type" in config:
            if config["type"] not in self.animation_types:
                return False
        
        if "duration" in config:
            duration = config["duration"]
            if isinstance(duration, str):
                if duration not in self.timing_options:
                    return False
            elif not isinstance(duration, (int, float)):
                return False
        
        if "trigger" in config:
            valid_triggers = ["on_click", "with_previous", "after_previous", "on_page_click"]
            if config["trigger"] not in valid_triggers:
                return False
        
        return True
    
    def validate_transition_config(self, config: Dict[str, Any]) -> bool:
        """Validate transition configuration."""
        if "type" in config:
            if config["type"] not in self.transition_types:
                return False
        
        if "duration" in config:
            duration = config["duration"]
            if isinstance(duration, str):
                if duration not in self.timing_options:
                    return False
            elif not isinstance(duration, (int, float)):
                return False
        
        return True


class TransitionPresets:
    """Predefined transition configurations."""
    
    @staticmethod
    def get_preset(preset_name: str) -> Optional[Dict[str, Any]]:
        """Get predefined transition configuration."""
        presets = {
            "smooth": {
                "type": "fade",
                "duration": "medium"
            },
            "dynamic": {
                "type": "zoom",
                "duration": "fast"
            },
            "professional": {
                "type": "wipe",
                "duration": "medium"
            },
            "creative": {
                "type": "cube",
                "duration": "slow"
            },
            "minimal": {
                "type": "fade",
                "duration": "fast"
            }
        }
        return presets.get(preset_name)


class AnimationPresets:
    """Predefined animation configurations."""
    
    @staticmethod
    def get_preset(preset_name: str) -> Optional[Dict[str, Any]]:
        """Get predefined animation configuration."""
        presets = {
            "slide_in": {
                "type": "fly_in",
                "duration": "medium",
                "trigger": "on_click"
            },
            "fade_in": {
                "type": "fade_in",
                "duration": "medium",
                "trigger": "on_click"
            },
            "zoom_in": {
                "type": "zoom",
                "duration": "fast",
                "trigger": "on_click"
            },
            "bounce_in": {
                "type": "bounce",
                "duration": "medium",
                "trigger": "on_click"
            },
            "spin_in": {
                "type": "swivel",
                "duration": "medium",
                "trigger": "on_click"
            },
            "auto_fade": {
                "type": "fade_in",
                "duration": "medium",
                "trigger": "after_previous",
                "delay": 0.5
            },
            "sequence_fade": {
                "type": "fade_in",
                "duration": "fast",
                "trigger": "after_previous",
                "delay": 0.3
            }
        }
        return presets.get(preset_name)
    
    @staticmethod
    def get_sequence_preset(preset_name: str) -> Optional[Dict[str, Any]]:
        """Get predefined animation sequence configuration."""
        presets = {
            "cascade": {
                "type": "sequential",
                "base_delay": 0,
                "delay_increment": 0.3,
                "animation": {
                    "type": "fade_in",
                    "duration": "fast",
                    "trigger": "after_previous"
                }
            },
            "simultaneous": {
                "type": "simultaneous",
                "base_delay": 0,
                "delay_increment": 0,
                "animation": {
                    "type": "zoom",
                    "duration": "medium",
                    "trigger": "on_click"
                }
            },
            "wave": {
                "type": "sequential",
                "base_delay": 0,
                "delay_increment": 0.2,
                "animation": {
                    "type": "fly_in",
                    "duration": "fast",
                    "trigger": "after_previous"
                }
            }
        }
        return presets.get(preset_name)
