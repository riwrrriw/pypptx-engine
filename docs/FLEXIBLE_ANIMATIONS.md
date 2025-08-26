# Flexible Animation System

The pypptx-engine now includes a powerful, flexible animation system that allows users to create custom animations with full control over timing, easing, and effects.

## Animation Categories

### 1. Motion Animations
Move shapes in any direction with customizable parameters:

```json
{
  "animation": {
    "type": "move_right",
    "duration": "medium",
    "trigger": "on_click",
    "delay": 0.5,
    "custom": {
      "from": "#ppt_x",
      "to": "#ppt_x+0.5",
      "attr": "ppt_x",
      "easing": "spline"
    }
  }
}
```

**Available Motion Types:**
- `move_left`, `move_right`, `move_up`, `move_down`, `move_diagonal`

### 2. Entrance Animations
Control how shapes appear on slides:

```json
{
  "animation": {
    "type": "fly_in_left",
    "duration": "fast",
    "trigger": "after_previous",
    "custom": {
      "preset_id": "2",
      "preset_subtype": "8",
      "transition": "in",
      "filter": "slide"
    }
  }
}
```

**Available Entrance Types:**
- `fade_in`, `fly_in_left`, `fly_in_right`, `fly_in_top`, `fly_in_bottom`, `zoom_in`, `bounce_in`

### 3. Emphasis Animations
Add emphasis effects to existing shapes:

```json
{
  "animation": {
    "type": "pulse",
    "duration": "medium",
    "trigger": "on_click",
    "repeat": 2,
    "custom": {
      "filter": "pulse",
      "transition": "emphasis"
    }
  }
}
```

**Available Emphasis Types:**
- `pulse`, `color_pulse`, `grow_shrink`, `spin`, `bounce`

## Customization Options

### Basic Parameters
- **`type`**: Animation effect name
- **`duration`**: `"very_fast"`, `"fast"`, `"medium"`, `"slow"`, `"very_slow"` or numeric seconds
- **`trigger`**: `"on_click"`, `"with_previous"`, `"after_previous"`, `"auto"`
- **`delay`**: Delay in seconds before animation starts
- **`repeat`**: Number of times to repeat animation
- **`reverse`**: Boolean to reverse animation direction
- **`easing`**: `"linear"`, `"spline"`, `"bounce"` for animation curves

### Advanced Custom Parameters
Use the `custom` object for fine-grained control:

#### Motion Customization
```json
{
  "custom": {
    "from": "#ppt_x",           // Starting position
    "to": "#ppt_x+0.75",        // Ending position  
    "attr": "ppt_x",            // Attribute to animate (ppt_x, ppt_y)
    "easing": "spline"          // Animation curve
  }
}
```

#### Entrance/Emphasis Customization
```json
{
  "custom": {
    "preset_id": "10",          // PowerPoint preset ID
    "preset_subtype": "4",      // Preset variation
    "transition": "in",         // Transition type
    "filter": "fade"            // Effect filter
  }
}
```

#### Complete XML Control
For maximum flexibility, provide your own XML template:

```json
{
  "custom": {
    "xml_template": "<p:timing>...{shape_id}...{duration}...{delay}...{trigger}...</p:timing>"
  }
}
```

## PowerPoint Expression System

Use PowerPoint's built-in expression system for dynamic positioning:

- **`#ppt_x`**: Current X position
- **`#ppt_y`**: Current Y position  
- **`#ppt_w`**: Shape width
- **`#ppt_h`**: Shape height
- **Mathematical operations**: `+`, `-`, `*`, `/`

### Examples:
- `"#ppt_x+0.25"`: Move 25% slide width to the right
- `"#ppt_y-#ppt_h"`: Move up by shape height
- `"0-#ppt_w/2"`: Move to left edge minus half width

## Trigger Events

Control when animations start:

- **`on_click`**: Start on mouse click/spacebar
- **`with_previous`**: Start simultaneously with previous animation
- **`after_previous`**: Start after previous animation completes
- **`auto`**: Start automatically after delay

## Complete Example

```json
{
  "type": "autoshape",
  "shape_type": "RECTANGLE",
  "text": "Advanced Animation",
  "x": 2, "y": 4, "w": 4, "h": 2,
  "animation": {
    "type": "move_right",
    "duration": 2.5,
    "trigger": "after_previous", 
    "delay": 0.3,
    "easing": "bounce",
    "repeat": 1,
    "reverse": false,
    "custom": {
      "from": "#ppt_x",
      "to": "#ppt_x+0.6", 
      "attr": "ppt_x",
      "easing": "spline"
    }
  }
}
```

## Animation Sequencing

Create complex animation sequences by combining different triggers:

1. **First shape**: `"trigger": "on_click"` - Starts on click
2. **Second shape**: `"trigger": "with_previous"` - Starts simultaneously  
3. **Third shape**: `"trigger": "after_previous", "delay": 0.5` - Starts 0.5s after second completes

This flexible system allows you to create professional PowerPoint animations with complete control over timing, movement, and effects while maintaining compatibility with PowerPoint's native animation engine.
