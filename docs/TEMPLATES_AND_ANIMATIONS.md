# Templates and Animations Guide

## Overview

The PPTX Engine now supports **templates** and **animations/transitions** to create professional, dynamic presentations with minimal JSON configuration.

## Template System

### Available Templates

1. **Corporate** - Professional business template
2. **Modern** - Clean contemporary design
3. **Creative** - Vibrant artistic template
4. **Academic** - Formal academic presentation

### Available Themes

- **Blue** - Professional blue color scheme
- **Green** - Nature-inspired green palette
- **Purple** - Creative purple theme
- **Orange** - Energetic orange theme

### Using Templates in JSON

```json
{
  "presentation": {
    "properties": {
      "title": "My Presentation",
      "template": "corporate",
      "theme": "blue"
    },
    "slides": [
      {
        "shapes": [
          {
            "type": "text",
            "text": "Title will use template styling"
          }
        ]
      }
    ]
  }
}
```

### Template Features

- **Automatic Styling**: Text shapes get styled based on position (title vs content)
- **Background Themes**: Gradient or solid backgrounds with theme colors
- **Font Consistency**: Template-specific font families and sizes
- **Position Guidelines**: Suggested positioning for optimal layout

## Animation System

### Slide Transitions

Add smooth transitions between slides:

```json
{
  "layout": 6,
  "transition": {
    "type": "fade",
    "duration": "medium"
  }
}
```

#### Available Transitions

- `fade`, `push`, `wipe`, `split`, `reveal`
- `zoom`, `rotate`, `cube`, `flip`, `flash`
- `pan`, `glitter`, `honeycomb`, `shred`

#### Duration Options

- `"very_fast"` (0.5s)
- `"fast"` (1.0s)
- `"medium"` (2.0s)
- `"slow"` (3.0s)
- `"very_slow"` (5.0s)
- Custom: `1.5` (numeric seconds)

### Shape Animations

Animate individual shapes:

```json
{
  "type": "text",
  "text": "Animated text",
  "animation": {
    "type": "fade_in",
    "duration": "medium",
    "trigger": "on_click",
    "delay": 0.5
  }
}
```

#### Animation Types

**Entrance Animations:**
- `appear`, `fade_in`, `fly_in`, `float_in`
- `zoom`, `bounce`, `swivel`, `grow_and_turn`
- `split`, `wipe`, `wheel`, `random_bars`

**Emphasis Animations:**
- `pulse`, `spin`, `grow_shrink`, `teeter`
- `color_pulse`, `transparency`, `darken`, `lighten`

**Exit Animations:**
- `disappear`, `fade_out`, `fly_out`, `float_out`
- `zoom_out`, `bounce_out`, `swivel_out`

#### Animation Triggers

- `"on_click"` - Start on mouse click
- `"with_previous"` - Start with previous animation
- `"after_previous"` - Start after previous animation ends
- `"on_page_click"` - Start when slide is clicked

### Animation Presets

Use predefined animation configurations:

```json
{
  "animation": "fade_in"
}
```

Available presets: `slide_in`, `fade_in`, `zoom_in`, `bounce_in`, `spin_in`, `auto_fade`, `sequence_fade`

### Transition Presets

```json
{
  "transition": "smooth"
}
```

Available presets: `smooth`, `dynamic`, `professional`, `creative`, `minimal`

## Complete Example

```json
{
  "presentation": {
    "properties": {
      "title": "Animated Corporate Presentation",
      "template": "corporate",
      "theme": "blue"
    },
    "slides": [
      {
        "transition": {
          "type": "fade",
          "duration": "medium"
        },
        "shapes": [
          {
            "type": "text",
            "text": "Welcome",
            "animation": {
              "type": "zoom",
              "duration": "fast",
              "trigger": "on_click"
            }
          },
          {
            "type": "text",
            "text": ["Point 1", "Point 2", "Point 3"],
            "animation": {
              "type": "fade_in",
              "duration": "medium",
              "trigger": "after_previous",
              "delay": 0.5
            }
          },
          {
            "type": "image",
            "url": "https://example.com/image.jpg",
            "x": 10, "y": 2, "w": 4, "h": 3,
            "animation": {
              "type": "float_in",
              "duration": "slow",
              "trigger": "after_previous"
            }
          }
        ]
      }
    ]
  }
}
```

## Advanced Features

### Template Customization

Templates automatically apply styling based on:
- **Shape position** (titles vs content)
- **Shape size** (determines hierarchy)
- **Theme colors** (primary, secondary, accent)

### Animation Sequences

Create cascading animations:

```json
{
  "shapes": [
    {
      "type": "text",
      "text": "First",
      "animation": {
        "type": "fade_in",
        "trigger": "on_click",
        "delay": 0
      }
    },
    {
      "type": "text", 
      "text": "Second",
      "animation": {
        "type": "fade_in",
        "trigger": "after_previous",
        "delay": 0.3
      }
    },
    {
      "type": "text",
      "text": "Third", 
      "animation": {
        "type": "fade_in",
        "trigger": "after_previous",
        "delay": 0.3
      }
    }
  ]
}
```

## Best Practices

### Templates
- Choose templates that match your presentation purpose
- Use consistent themes across slides
- Let templates handle positioning for optimal layouts

### Animations
- Use entrance animations for key points
- Keep durations consistent (medium is usually best)
- Use `after_previous` with delays for smooth sequences
- Avoid overusing animations - less is more

### Performance
- Templates reduce JSON configuration size
- Animations are stored as metadata (limited by python-pptx)
- Use presets for consistency and simplicity

## Troubleshooting

- **Template not applying**: Check template name spelling
- **Animation not working**: Verify animation type is supported
- **Timing issues**: Adjust delay values for better flow
- **Theme colors not showing**: Ensure theme name is correct

## Examples Directory

Check these example files:
- `template_corporate.json` - Corporate template with animations
- `template_modern.json` - Modern template with flowcharts
- `template_creative.json` - Creative template with effects
