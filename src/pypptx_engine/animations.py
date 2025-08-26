"""
Animation and transition system for PPTX Engine.
Provides slide transitions and shape animations.
"""

from typing import Dict, Any, List, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
import warnings


class AnimationManager:
    """Manages animations for PowerPoint presentations with flexible customization."""
    
    def __init__(self):
        """Initialize the animation manager."""
        self.timing_options = {
            "very_fast": 0.25,
            "fast": 0.5,
            "medium": 1.0,
            "slow": 2.0,
            "very_slow": 3.0
        }
        
        # Animation presets for different effects
        self.animation_presets = {
            "entrance": {
                "fade_in": {"preset_id": "1", "preset_class": "entr", "preset_subtype": "0"},
                "fly_in_left": {"preset_id": "2", "preset_class": "entr", "preset_subtype": "8"},
                "fly_in_right": {"preset_id": "2", "preset_class": "entr", "preset_subtype": "2"},
                "fly_in_top": {"preset_id": "2", "preset_class": "entr", "preset_subtype": "4"},
                "fly_in_bottom": {"preset_id": "2", "preset_class": "entr", "preset_subtype": "6"},
                "zoom_in": {"preset_id": "10", "preset_class": "entr", "preset_subtype": "0"},
                "bounce_in": {"preset_id": "26", "preset_class": "entr", "preset_subtype": "0"}
            },
            "emphasis": {
                "pulse": {"preset_id": "1", "preset_class": "emph", "preset_subtype": "0"},
                "color_pulse": {"preset_id": "2", "preset_class": "emph", "preset_subtype": "0"},
                "grow_shrink": {"preset_id": "3", "preset_class": "emph", "preset_subtype": "0"},
                "spin": {"preset_id": "5", "preset_class": "emph", "preset_subtype": "0"},
                "bounce": {"preset_id": "26", "preset_class": "emph", "preset_subtype": "0"}
            },
            "motion": {
                "move_left": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x-0.25"},
                "move_right": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x+0.25"},
                "move_up": {"attr": "ppt_y", "from": "#ppt_y", "to": "#ppt_y-0.25"},
                "move_down": {"attr": "ppt_y", "from": "#ppt_y", "to": "#ppt_y+0.25"},
                "move_diagonal": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x+0.25"},
                "custom_path": {"attr": "custom", "from": "custom", "to": "custom"}
            }
        }
        
        # Trigger event mappings
        self.trigger_events = {
            "on_click": "onNext",
            "with_previous": "withPrev",
            "after_previous": "afterPrev",
            "on_page_click": "onNext",
            "auto": "afterPrev"
        }
        
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
        """Apply transition effect to slide using XML injection."""
        try:
            from pptx.oxml import parse_xml
            
            transition_type = transition_config.get("type", "none")
            duration = transition_config.get("duration", 1.0)
            
            # Convert string duration to float if needed
            if isinstance(duration, str):
                duration = self.timing_options.get(duration, 1.0)
            
            # Convert duration to milliseconds for PowerPoint
            duration_ms = int(duration * 1000)
            
            # Get XML template for transition
            transition_xml = self._get_transition_xml(transition_type, duration_ms)
            
            if transition_xml:
                # Parse and inject XML into slide
                xml_fragment = parse_xml(transition_xml)
                slide.element.insert(-1, xml_fragment)
                    
        except Exception as e:
            warnings.warn(f"Failed to apply slide transition: {e}")
    
    def _get_transition_xml(self, transition_type: str, duration_ms: int) -> str:
        """Get XML template for specific transition type."""
        # Speed mapping for PowerPoint
        if duration_ms <= 500:
            speed = "fast"
        elif duration_ms <= 1500:
            speed = "med"
        else:
            speed = "slow"
        
        transition_templates = {
            "fade": f'''
                <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                    <p:fade />
                </p:transition>
            ''',
            "push": f'''
                <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                    <p:push dir="l" />
                </p:transition>
            ''',
            "wipe": f'''
                <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                    <p:wipe dir="l" />
                </p:transition>
            ''',
            "split": f'''
                <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                    <p:split orient="horz" dir="out" />
                </p:transition>
            ''',
            "reveal": f'''
                <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                    <p:reveal dir="l" />
                </p:transition>
            ''',
            "zoom": f'''
                <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}">
                            <p14:zoom />
                        </p:transition>
                    </mc:Choice>
                    <mc:Fallback>
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}">
                            <p:fade />
                        </p:transition>
                    </mc:Fallback>
                </mc:AlternateContent>
            ''',
            "cube": f'''
                <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}">
                            <p14:prism />
                        </p:transition>
                    </mc:Choice>
                    <mc:Fallback>
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}">
                            <p:fade />
                        </p:transition>
                    </mc:Fallback>
                </mc:AlternateContent>
            ''',
            "flip": f'''
                <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}">
                            <p14:flip />
                        </p:transition>
                    </mc:Choice>
                    <mc:Fallback>
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}">
                            <p:fade />
                        </p:transition>
                    </mc:Fallback>
                </mc:AlternateContent>
            ''',
            "rotate": f'''
                <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}" p14:dur="{duration_ms}">
                            <p14:doors />
                        </p:transition>
                    </mc:Choice>
                    <mc:Fallback>
                        <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="{speed}">
                            <p:wipe dir="l" />
                        </p:transition>
                    </mc:Fallback>
                </mc:AlternateContent>
            '''
        }
        
        return transition_templates.get(transition_type, "")
    
    def apply_shape_animation(self, slide, shape, animation_config: Dict[str, Any]) -> None:
        """Apply flexible animation to shape with customizable options."""
        try:
            # Parse animation configuration
            animation_type = animation_config.get("type", "fade_in")
            delay = animation_config.get("delay", 0)
            duration = animation_config.get("duration", "medium")
            trigger = animation_config.get("trigger", "on_click")
            
            # Advanced customization options
            custom_params = animation_config.get("custom", {})
            easing = animation_config.get("easing", "linear")
            repeat = animation_config.get("repeat", 1)
            reverse = animation_config.get("reverse", False)
            
            # Convert string duration to float
            if isinstance(duration, str):
                duration = self.timing_options.get(duration, 1.0)
            
            # Build flexible animation based on type
            self._build_flexible_animation(slide, shape, {
                "type": animation_type,
                "duration": duration,
                "delay": delay,
                "trigger": trigger,
                "custom": custom_params,
                "easing": easing,
                "repeat": repeat,
                "reverse": reverse
            })
                    
        except Exception as e:
            warnings.warn(f"Failed to apply shape animation: {e}")
    
    def _inject_motion_animation_directly(self, slide, shape, animation_type: str, duration: float, delay: float, trigger: str):
        """Inject motion animation using actual PowerPoint XML structure from extracted file."""
        try:
            from pptx.oxml import parse_xml
            
            # Get shape ID
            shape_id = shape.shape_id if hasattr(shape, 'shape_id') else 2
            duration_ms = int(duration * 1000)
            delay_ms = int(delay * 1000)
            
            # Define motion animation values based on direction
            direction = animation_type.replace("move_", "")
            animation_values = {
                "left": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x-0.25"},
                "right": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x+0.25"},
                "up": {"attr": "ppt_y", "from": "#ppt_y", "to": "#ppt_y-0.25"},
                "down": {"attr": "ppt_y", "from": "#ppt_y", "to": "#ppt_y+0.25"},
                "diagonal": {"attr": "ppt_x", "from": "#ppt_x", "to": "#ppt_x+0.25"}
            }
            
            anim_config = animation_values.get(direction, animation_values["right"])
            
            # Create timing XML based on actual PowerPoint structure from extracted file
            timing_xml = f'''
                <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                    <p:tnLst>
                        <p:par>
                            <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                                <p:childTnLst>
                                    <p:seq concurrent="1" nextAc="seek">
                                        <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                            <p:childTnLst>
                                                <p:par>
                                                    <p:cTn id="3" fill="hold">
                                                        <p:stCondLst>
                                                            <p:cond delay="indefinite"/>
                                                        </p:stCondLst>
                                                        <p:childTnLst>
                                                            <p:par>
                                                                <p:cTn id="4" fill="hold">
                                                                    <p:stCondLst>
                                                                        <p:cond delay="0"/>
                                                                    </p:stCondLst>
                                                                    <p:childTnLst>
                                                                        <p:par>
                                                                            <p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="8" fill="hold" grpId="0" nodeType="clickEffect">
                                                                                <p:stCondLst>
                                                                                    <p:cond delay="0"/>
                                                                                </p:stCondLst>
                                                                                <p:childTnLst>
                                                                                    <p:set>
                                                                                        <p:cBhvr>
                                                                                            <p:cTn id="6" dur="1" fill="hold">
                                                                                                <p:stCondLst>
                                                                                                    <p:cond delay="0"/>
                                                                                                </p:stCondLst>
                                                                                            </p:cTn>
                                                                                            <p:tgtEl>
                                                                                                <p:spTgt spid="{shape_id}"/>
                                                                                            </p:tgtEl>
                                                                                            <p:attrNameLst>
                                                                                                <p:attrName>style.visibility</p:attrName>
                                                                                            </p:attrNameLst>
                                                                                        </p:cBhvr>
                                                                                        <p:to>
                                                                                            <p:strVal val="visible"/>
                                                                                        </p:to>
                                                                                    </p:set>
                                                                                    <p:anim calcmode="lin" valueType="num">
                                                                                        <p:cBhvr additive="base">
                                                                                            <p:cTn id="7" dur="500" fill="hold"/>
                                                                                            <p:tgtEl>
                                                                                                <p:spTgt spid="{shape_id}"/>
                                                                                            </p:tgtEl>
                                                                                            <p:attrNameLst>
                                                                                                <p:attrName>ppt_x</p:attrName>
                                                                                            </p:attrNameLst>
                                                                                        </p:cBhvr>
                                                                                        <p:tavLst>
                                                                                            <p:tav tm="0">
                                                                                                <p:val>
                                                                                                    <p:strVal val="0-#ppt_w/2"/>
                                                                                                </p:val>
                                                                                            </p:tav>
                                                                                            <p:tav tm="100000">
                                                                                                <p:val>
                                                                                                    <p:strVal val="#ppt_x"/>
                                                                                                </p:val>
                                                                                            </p:tav>
                                                                                        </p:tavLst>
                                                                                    </p:anim>
                                                                                </p:childTnLst>
                                                                            </p:cTn>
                                                                        </p:par>
                                                                    </p:childTnLst>
                                                                </p:cTn>
                                                            </p:par>
                                                        </p:childTnLst>
                                                    </p:cTn>
                                                </p:par>
                                            </p:childTnLst>
                                        </p:cTn>
                                        <p:prevCondLst>
                                            <p:cond evt="onPrev" delay="0">
                                                <p:tgtEl>
                                                    <p:sldTgt/>
                                                </p:tgtEl>
                                            </p:cond>
                                        </p:prevCondLst>
                                        <p:nextCondLst>
                                            <p:cond evt="onNext" delay="0">
                                                <p:tgtEl>
                                                    <p:sldTgt/>
                                                </p:tgtEl>
                                            </p:cond>
                                        </p:nextCondLst>
                                    </p:seq>
                                    <p:seq concurrent="1" nextAc="seek">
                                        <p:cTn id="9" dur="indefinite" nodeType="mainSeq">
                                            <p:childTnLst>
                                                <p:par>
                                                    <p:cTn id="10" fill="hold">
                                                        <p:stCondLst>
                                                            <p:cond evt="onNext" delay="{delay_ms}"/>
                                                        </p:stCondLst>
                                                        <p:childTnLst>
                                                            <p:anim from="{anim_config['from']}" to="{anim_config['to']}" calcmode="lin" valueType="num">
                                                                <p:cBhvr override="childStyle">
                                                                    <p:cTn id="11" dur="{duration_ms}" fill="hold"/>
                                                                    <p:tgtEl>
                                                                        <p:spTgt spid="{shape_id}"/>
                                                                    </p:tgtEl>
                                                                    <p:attrNameLst>
                                                                        <p:attrName>{anim_config['attr']}</p:attrName>
                                                                    </p:attrNameLst>
                                                                </p:cBhvr>
                                                            </p:anim>
                                                        </p:childTnLst>
                                                    </p:cTn>
                                                </p:par>
                                            </p:childTnLst>
                                        </p:cTn>
                                    </p:seq>
                                </p:childTnLst>
                            </p:cTn>
                        </p:par>
                    </p:tnLst>
                    <p:bldLst>
                        <p:bldP spid="{shape_id}" grpId="0" animBg="1"/>
                    </p:bldLst>
                </p:timing>
            '''
            
            # Remove existing timing element if present
            existing_timing = slide.element.xpath('.//p:timing')
            for timing in existing_timing:
                timing.getparent().remove(timing)
            
            # Add new timing element
            timing_fragment = parse_xml(timing_xml)
            slide.element.append(timing_fragment)
            
        except Exception as e:
            warnings.warn(f"Failed to inject motion animation: {e}")
    
    def _build_flexible_animation(self, slide, shape, config: Dict[str, Any]) -> None:
        """Build flexible animation with customizable parameters."""
        try:
            from pptx.oxml import parse_xml
            
            animation_type = config["type"]
            duration_ms = int(config["duration"] * 1000)
            delay_ms = int(config["delay"] * 1000)
            trigger = self.trigger_events.get(config["trigger"], "onNext")
            custom_params = config.get("custom", {})
            
            shape_id = shape.shape_id if hasattr(shape, 'shape_id') else 2
            
            # Determine animation category and build accordingly
            if animation_type.startswith("move_") or animation_type in self.animation_presets["motion"]:
                self._build_motion_animation(slide, shape_id, animation_type, duration_ms, delay_ms, trigger, custom_params)
            elif animation_type in self.animation_presets["entrance"]:
                self._build_entrance_animation(slide, shape_id, animation_type, duration_ms, delay_ms, trigger, custom_params)
            elif animation_type in self.animation_presets["emphasis"]:
                self._build_emphasis_animation(slide, shape_id, animation_type, duration_ms, delay_ms, trigger, custom_params)
            else:
                # Custom animation with user-defined parameters
                self._build_custom_animation(slide, shape_id, config)
                
        except Exception as e:
            warnings.warn(f"Failed to build flexible animation: {e}")
    
    def _build_motion_animation(self, slide, shape_id: int, animation_type: str, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> None:
        """Build customizable motion animation."""
        try:
            from pptx.oxml import parse_xml
            
            # Get motion configuration
            direction = animation_type.replace("move_", "")
            
            # Allow custom motion values
            if "from" in custom_params and "to" in custom_params:
                motion_config = {
                    "attr": custom_params.get("attr", "ppt_x"),
                    "from": custom_params["from"],
                    "to": custom_params["to"]
                }
            else:
                motion_config = self.animation_presets["motion"].get(animation_type, 
                    self.animation_presets["motion"]["move_right"])
            
            # Build timing XML with custom parameters
            timing_xml = self._generate_motion_timing_xml(shape_id, motion_config, duration_ms, delay_ms, trigger, custom_params)
            
            # Apply to slide
            self._apply_timing_xml(slide, timing_xml)
            
        except Exception as e:
            warnings.warn(f"Failed to build motion animation: {e}")
    
    def _build_entrance_animation(self, slide, shape_id: int, animation_type: str, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> None:
        """Build customizable entrance animation."""
        try:
            from pptx.oxml import parse_xml
            
            preset = self.animation_presets["entrance"].get(animation_type, 
                self.animation_presets["entrance"]["fade_in"])
            
            # Allow custom preset overrides
            if "preset_id" in custom_params:
                preset["preset_id"] = custom_params["preset_id"]
            if "preset_subtype" in custom_params:
                preset["preset_subtype"] = custom_params["preset_subtype"]
            
            timing_xml = self._generate_entrance_timing_xml(shape_id, preset, duration_ms, delay_ms, trigger, custom_params)
            self._apply_timing_xml(slide, timing_xml)
            
        except Exception as e:
            warnings.warn(f"Failed to build entrance animation: {e}")
    
    def _build_emphasis_animation(self, slide, shape_id: int, animation_type: str, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> None:
        """Build customizable emphasis animation."""
        try:
            from pptx.oxml import parse_xml
            
            preset = self.animation_presets["emphasis"].get(animation_type, 
                self.animation_presets["emphasis"]["pulse"])
            
            timing_xml = self._generate_emphasis_timing_xml(shape_id, preset, duration_ms, delay_ms, trigger, custom_params)
            self._apply_timing_xml(slide, timing_xml)
            
        except Exception as e:
            warnings.warn(f"Failed to build emphasis animation: {e}")
    
    def _build_custom_animation(self, slide, shape_id: int, config: Dict[str, Any]) -> None:
        """Build completely custom animation from user parameters."""
        try:
            from pptx.oxml import parse_xml
            
            custom_params = config.get("custom", {})
            
            # User can define complete custom XML or use helper parameters
            if "xml_template" in custom_params:
                # Use user-provided XML template
                timing_xml = custom_params["xml_template"].format(
                    shape_id=shape_id,
                    duration=int(config["duration"] * 1000),
                    delay=int(config["delay"] * 1000),
                    trigger=self.trigger_events.get(config["trigger"], "onNext")
                )
            else:
                # Build from custom parameters
                timing_xml = self._generate_custom_timing_xml(shape_id, config)
            
            self._apply_timing_xml(slide, timing_xml)
            
        except Exception as e:
            warnings.warn(f"Failed to build custom animation: {e}")
    
    def _apply_timing_xml(self, slide, timing_xml: str) -> None:
        """Apply timing XML to slide."""
        try:
            from pptx.oxml import parse_xml
            
            # Remove existing timing element if present
            existing_timing = slide.element.xpath('.//p:timing')
            for timing in existing_timing:
                timing.getparent().remove(timing)
            
            # Add new timing element
            timing_fragment = parse_xml(timing_xml)
            slide.element.append(timing_fragment)
            
        except Exception as e:
            warnings.warn(f"Failed to apply timing XML: {e}")
    
    def _generate_motion_timing_xml(self, shape_id: int, motion_config: Dict, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> str:
        """Generate motion timing XML with customizable parameters."""
        return f'''
            <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:tnLst>
                    <p:par>
                        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                            <p:childTnLst>
                                <p:seq concurrent="1" nextAc="seek">
                                    <p:cTn id="9" dur="indefinite" nodeType="mainSeq">
                                        <p:childTnLst>
                                            <p:par>
                                                <p:cTn id="10" fill="hold">
                                                    <p:stCondLst>
                                                        <p:cond evt="{trigger}" delay="{delay_ms}"/>
                                                    </p:stCondLst>
                                                    <p:childTnLst>
                                                        <p:anim from="{motion_config['from']}" to="{motion_config['to']}" calcmode="{custom_params.get('easing', 'lin')}" valueType="num">
                                                            <p:cBhvr override="childStyle">
                                                                <p:cTn id="11" dur="{duration_ms}" fill="hold"/>
                                                                <p:tgtEl>
                                                                    <p:spTgt spid="{shape_id}"/>
                                                                </p:tgtEl>
                                                                <p:attrNameLst>
                                                                    <p:attrName>{motion_config['attr']}</p:attrName>
                                                                </p:attrNameLst>
                                                            </p:cBhvr>
                                                        </p:anim>
                                                    </p:childTnLst>
                                                </p:cTn>
                                            </p:par>
                                        </p:childTnLst>
                                    </p:cTn>
                                </p:seq>
                            </p:childTnLst>
                        </p:cTn>
                    </p:par>
                </p:tnLst>
                <p:bldLst>
                    <p:bldP spid="{shape_id}" grpId="0" animBg="1"/>
                </p:bldLst>
            </p:timing>
        '''
    
    def _generate_entrance_timing_xml(self, shape_id: int, preset: Dict, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> str:
        """Generate entrance timing XML with customizable parameters."""
        return f'''
            <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:tnLst>
                    <p:par>
                        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                            <p:childTnLst>
                                <p:seq concurrent="1" nextAc="seek">
                                    <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                        <p:childTnLst>
                                            <p:par>
                                                <p:cTn id="5" presetID="{preset['preset_id']}" presetClass="{preset['preset_class']}" presetSubtype="{preset['preset_subtype']}" fill="hold" grpId="0" nodeType="clickEffect">
                                                    <p:stCondLst>
                                                        <p:cond evt="{trigger}" delay="{delay_ms}"/>
                                                    </p:stCondLst>
                                                    <p:childTnLst>
                                                        <p:set>
                                                            <p:cBhvr>
                                                                <p:cTn id="6" dur="1" fill="hold"/>
                                                                <p:tgtEl>
                                                                    <p:spTgt spid="{shape_id}"/>
                                                                </p:tgtEl>
                                                                <p:attrNameLst>
                                                                    <p:attrName>style.visibility</p:attrName>
                                                                </p:attrNameLst>
                                                            </p:cBhvr>
                                                            <p:to>
                                                                <p:strVal val="visible"/>
                                                            </p:to>
                                                        </p:set>
                                                        <p:animEffect transition="{custom_params.get('transition', 'in')}" filter="{custom_params.get('filter', 'fade')}">
                                                            <p:cBhvr>
                                                                <p:cTn id="7" dur="{duration_ms}" fill="hold"/>
                                                                <p:tgtEl>
                                                                    <p:spTgt spid="{shape_id}"/>
                                                                </p:tgtEl>
                                                            </p:cBhvr>
                                                        </p:animEffect>
                                                    </p:childTnLst>
                                                </p:cTn>
                                            </p:par>
                                        </p:childTnLst>
                                    </p:cTn>
                                </p:seq>
                            </p:childTnLst>
                        </p:cTn>
                    </p:par>
                </p:tnLst>
                <p:bldLst>
                    <p:bldP spid="{shape_id}" grpId="0" animBg="1"/>
                </p:bldLst>
            </p:timing>
        '''
    
    def _generate_emphasis_timing_xml(self, shape_id: int, preset: Dict, duration_ms: int, delay_ms: int, trigger: str, custom_params: Dict) -> str:
        """Generate emphasis timing XML with customizable parameters."""
        return f'''
            <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:tnLst>
                    <p:par>
                        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                            <p:childTnLst>
                                <p:seq concurrent="1" nextAc="seek">
                                    <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                        <p:childTnLst>
                                            <p:par>
                                                <p:cTn id="5" presetID="{preset['preset_id']}" presetClass="{preset['preset_class']}" presetSubtype="{preset['preset_subtype']}" fill="hold" grpId="0" nodeType="clickEffect">
                                                    <p:stCondLst>
                                                        <p:cond evt="{trigger}" delay="{delay_ms}"/>
                                                    </p:stCondLst>
                                                    <p:childTnLst>
                                                        <p:animEffect transition="{custom_params.get('transition', 'in')}" filter="{custom_params.get('filter', 'emphasis')}">
                                                            <p:cBhvr>
                                                                <p:cTn id="7" dur="{duration_ms}" fill="hold"/>
                                                                <p:tgtEl>
                                                                    <p:spTgt spid="{shape_id}"/>
                                                                </p:tgtEl>
                                                            </p:cBhvr>
                                                        </p:animEffect>
                                                    </p:childTnLst>
                                                </p:cTn>
                                            </p:par>
                                        </p:childTnLst>
                                    </p:cTn>
                                </p:seq>
                            </p:childTnLst>
                        </p:cTn>
                    </p:par>
                </p:tnLst>
                <p:bldLst>
                    <p:bldP spid="{shape_id}" grpId="0" animBg="1"/>
                </p:bldLst>
            </p:timing>
        '''
    
    def _generate_custom_timing_xml(self, shape_id: int, config: Dict[str, Any]) -> str:
        """Generate completely custom timing XML from user configuration."""
        custom_params = config.get('custom', {})
        duration_ms = int(config['duration'] * 1000)
        delay_ms = int(config['delay'] * 1000)
        trigger = self.trigger_events.get(config['trigger'], 'onNext')
        
        # Default custom animation template
        return f'''
            <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:tnLst>
                    <p:par>
                        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                            <p:childTnLst>
                                <p:seq concurrent="1" nextAc="seek">
                                    <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                        <p:childTnLst>
                                            <p:par>
                                                <p:cTn id="5" fill="hold" grpId="0" nodeType="clickEffect">
                                                    <p:stCondLst>
                                                        <p:cond evt="{trigger}" delay="{delay_ms}"/>
                                                    </p:stCondLst>
                                                    <p:childTnLst>
                                                        <p:anim from="{custom_params.get('from', '#ppt_x')}" to="{custom_params.get('to', '#ppt_x+0.1')}" calcmode="{custom_params.get('easing', 'lin')}" valueType="num">
                                                            <p:cBhvr>
                                                                <p:cTn id="6" dur="{duration_ms}" fill="hold"/>
                                                                <p:tgtEl>
                                                                    <p:spTgt spid="{shape_id}"/>
                                                                </p:tgtEl>
                                                                <p:attrNameLst>
                                                                    <p:attrName>{custom_params.get('attr', 'ppt_x')}</p:attrName>
                                                                </p:attrNameLst>
                                                            </p:cBhvr>
                                                        </p:anim>
                                                    </p:childTnLst>
                                                </p:cTn>
                                            </p:par>
                                        </p:childTnLst>
                                    </p:cTn>
                                </p:seq>
                            </p:childTnLst>
                        </p:cTn>
                    </p:par>
                </p:tnLst>
                <p:bldLst>
                    <p:bldP spid="{shape_id}" grpId="0" animBg="1"/>
                </p:bldLst>
            </p:timing>
        '''
    
    def _get_or_create_timing_element(self, slide):
        """Get or create timing element in slide for animations."""
        try:
            from pptx.oxml import parse_xml
            
            # Check if timing element exists
            timing_elements = slide.element.xpath('.//p:timing')
            if timing_elements:
                # Get the main sequence container
                tnlst = timing_elements[0].xpath('.//p:tnLst')[0]
                main_par = tnlst.xpath('.//p:par')[0]
                childtnlst = main_par.xpath('.//p:childTnLst')
                if childtnlst:
                    return childtnlst[0]
                else:
                    # Create childTnLst if it doesn't exist
                    child_xml = '<p:childTnLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
                    child_fragment = parse_xml(child_xml)
                    main_par.append(child_fragment)
                    return child_fragment
            
            # Create complete timing structure if it doesn't exist
            timing_xml = '''
                <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                    <p:tnLst>
                        <p:par>
                            <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/>
                            <p:childTnLst>
                            </p:childTnLst>
                        </p:par>
                    </p:tnLst>
                </p:timing>
            '''
            timing_fragment = parse_xml(timing_xml)
            slide.element.append(timing_fragment)
            
            return slide.element.xpath('.//p:childTnLst')[0]
            
        except Exception as e:
            warnings.warn(f"Failed to create timing element: {e}")
            return None
    
    def _get_shape_animation_xml(self, shape, animation_type: str, duration: float, delay: float, trigger: str) -> str:
        """Get XML for shape animation."""
        # Convert duration to milliseconds
        duration_ms = int(duration * 1000)
        delay_ms = int(delay * 1000)
        
        # Get shape ID for targeting
        shape_id = shape.shape_id if hasattr(shape, 'shape_id') else 1
        
        # Map trigger types
        trigger_map = {
            "on_click": "onClick",
            "with_previous": "withPrev", 
            "after_previous": "afterPrev",
            "on_page_click": "onClick"
        }
        trigger_type = trigger_map.get(trigger, "onClick")
        
        # Animation effect mapping
        animation_effects = {
            "fade_in": {"type": "fade", "subtype": "none"},
            "fly_in": {"type": "fly", "subtype": "left"},
            "zoom": {"type": "zoom", "subtype": "in"},
            "bounce": {"type": "bounce", "subtype": "none"},
            "swivel": {"type": "swivel", "subtype": "none"},
            "appear": {"type": "appear", "subtype": "none"},
            "float_in": {"type": "float", "subtype": "up"},
            "grow_and_turn": {"type": "growTurn", "subtype": "none"},
            "spin": {"type": "spin", "subtype": "none"},
            "move_left": {"type": "path", "subtype": "left"},
            "move_right": {"type": "path", "subtype": "right"},
            "move_up": {"type": "path", "subtype": "up"},
            "move_down": {"type": "path", "subtype": "down"},
            "move_diagonal": {"type": "path", "subtype": "diagonal"}
        }
        
        effect = animation_effects.get(animation_type, {"type": "fade", "subtype": "none"})
        
        # Handle motion path animations differently
        if effect["type"] == "path":
            animation_xml = self._get_motion_path_xml(shape_id, effect["subtype"], duration_ms, delay_ms, trigger_type)
        else:
            animation_xml = f'''
                <p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                    <p:cTn id="{shape_id + 100}" fill="hold">
                        <p:stCondLst>
                            <p:cond evt="{trigger_type}" delay="{delay_ms}"/>
                        </p:stCondLst>
                        <p:childTnLst>
                            <p:animEffect transition="in" filter="{effect['type']}">
                                <p:cTn id="{shape_id + 200}" dur="{duration_ms}"/>
                                <p:tgtEl>
                                    <p:spTgt spid="{shape_id}"/>
                                </p:tgtEl>
                                <p:animBhv>
                                    <p:cTn id="{shape_id + 300}" dur="{duration_ms}"/>
                                    <p:tgtEl>
                                        <p:spTgt spid="{shape_id}"/>
                                    </p:tgtEl>
                                    <p:attrNameLst>
                                        <p:attrName>style.visibility</p:attrName>
                                    </p:attrNameLst>
                                </p:animBhv>
                            </p:animEffect>
                        </p:childTnLst>
                    </p:cTn>
                </p:par>
            '''
        
        return animation_xml
    
    def _get_motion_path_xml(self, shape_id: int, direction: str, duration_ms: int, delay_ms: int, trigger_type: str) -> str:
        """Generate XML for motion path animations."""
        # Use absolute coordinates for more reliable movement
        motion_paths = {
            "left": "M 0 0 L -100 0 E",
            "right": "M 0 0 L 100 0 E", 
            "up": "M 0 0 L 0 -100 E",
            "down": "M 0 0 L 0 100 E",
            "diagonal": "M 0 0 L 100 100 E"
        }
        
        path = motion_paths.get(direction, "M 0 0 L 100 0 E")
        
        return f'''
            <p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:cTn id="{shape_id + 100}" fill="hold">
                    <p:stCondLst>
                        <p:cond evt="{trigger_type}" delay="{delay_ms}"/>
                    </p:stCondLst>
                    <p:childTnLst>
                        <p:animMotion origin="layout" path="{path}" pathEditMode="fixed" ptsTypes="">
                            <p:cBhvr>
                                <p:cTn id="{shape_id + 200}" dur="{duration_ms}" fill="hold"/>
                                <p:tgtEl>
                                    <p:spTgt spid="{shape_id}"/>
                                </p:tgtEl>
                            </p:cBhvr>
                        </p:animMotion>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        '''
    
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
