# Flowchart Support in pypptx-engine

The pypptx-engine now supports comprehensive flowchart creation with automatic connection management and professional styling.

## Features

- **20+ Standard Flowchart Shapes**: Start/End, Process, Decision, Data, Document, and many more
- **Automatic Connection Management**: Smart routing between flowchart elements
- **Customizable Styling**: Predefined themes with full customization support
- **Connection Labels**: Add text labels to connections with positioning
- **Multiple Connector Types**: Straight, elbow, and curved connectors
- **Layout Utilities**: Automatic positioning helpers for common layouts
- **Professional Themes**: Built-in color schemes and formatting

## Supported Flowchart Shapes

| Shape Type | Description | JSON Key |
|------------|-------------|----------|
| Start/End | Oval shapes for process start/end | `start`, `end` |
| Process | Rectangle for process steps | `process` |
| Decision | Diamond for decision points | `decision` |
| Data/Input | Parallelogram for data input/output | `data` |
| Document | Document shape for reports/files | `document` |
| Manual Operation | Trapezoid for manual tasks | `manual_operation` |
| Predefined Process | Rectangle with double borders | `predefined_process` |
| Connector | Circle for connection points | `connector` |
| Stored Data | Cylinder for databases/storage | `stored_data` |
| Delay | D-shape for delays/waiting | `delay` |
| Display | Monitor shape for displays | `display` |
| Preparation | Hexagon for preparation steps | `preparation` |

## Basic Usage

### Simple Flowchart

```json
{
  "type": "flowchart",
  "elements": [
    {
      "id": "start",
      "flowchart_type": "start",
      "text": "Start",
      "x": 7, "y": 2, "w": 2, "h": 1
    },
    {
      "id": "process1",
      "flowchart_type": "process", 
      "text": "Process Data",
      "x": 6.5, "y": 4, "w": 3, "h": 1
    },
    {
      "id": "end",
      "flowchart_type": "end",
      "text": "End",
      "x": 7, "y": 6, "w": 2, "h": 1
    }
  ],
  "connections": [
    {
      "from": "start",
      "to": "process1",
      "connector_type": "STRAIGHT"
    },
    {
      "from": "process1", 
      "to": "end",
      "connector_type": "STRAIGHT"
    }
  ]
}
```

### Decision Flow with Branches

```json
{
  "type": "flowchart",
  "elements": [
    {
      "id": "decision",
      "flowchart_type": "decision",
      "text": "Valid Input?",
      "x": 6.5, "y": 3.5, "w": 3, "h": 1.5
    },
    {
      "id": "process_yes",
      "flowchart_type": "process",
      "text": "Process Data",
      "x": 3, "y": 6, "w": 3, "h": 1
    },
    {
      "id": "process_no",
      "flowchart_type": "process", 
      "text": "Show Error",
      "x": 10, "y": 6, "w": 3, "h": 1
    }
  ],
  "connections": [
    {
      "from": "decision",
      "to": "process_yes",
      "connector_type": "STRAIGHT",
      "from_side": "bottom-left",
      "to_side": "top",
      "label": "Yes",
      "label_config": {
        "font": {"size": 10, "color": "#27ae60", "bold": true}
      }
    },
    {
      "from": "decision",
      "to": "process_no", 
      "connector_type": "STRAIGHT",
      "from_side": "bottom-right",
      "to_side": "top",
      "label": "No",
      "label_config": {
        "font": {"size": 10, "color": "#e74c3c", "bold": true}
      }
    }
  ]
}
```

## Connection Configuration

### Connection Points

Specify where connections attach to shapes:

- `top`, `bottom`, `left`, `right` - Center of each side
- `top-left`, `top-right`, `bottom-left`, `bottom-right` - Corners

### Connector Types

- `STRAIGHT` - Direct line between points
- `ELBOW` - Right-angled connector with bends
- `CURVED` - Curved connector (limited support)

### Connection Labels

Add text labels to connections:

```json
{
  "from": "decision",
  "to": "process",
  "label": "Yes",
  "label_config": {
    "font": {"size": 10, "color": "#27ae60", "bold": true},
    "background": {"type": "solid", "color": "#ffffff"},
    "w": 1, "h": 0.3
  }
}
```

## Styling and Customization

### Default Themes

Each flowchart shape type has predefined styling:

- **Start/End**: Green oval with white text
- **Process**: Blue rectangle with white text  
- **Decision**: Orange diamond with white text
- **Data**: Purple parallelogram with white text
- **Document**: Teal document shape with white text

### Custom Styling

Override default styles for any element:

```json
{
  "id": "custom_process",
  "flowchart_type": "process",
  "text": "Custom Process",
  "fill": {
    "type": "gradient",
    "stops": [
      {"color": "#667eea", "position": 0},
      {"color": "#764ba2", "position": 1}
    ]
  },
  "line": {"color": "#2c3e50", "width": 3},
  "font": {"color": "#ffffff", "bold": true, "size": 14},
  "shadow": {"visible": true, "color": "#000000"}
}
```

## Layout Utilities

Use the `FlowchartLayoutManager` for automatic positioning:

```python
from pypptx_engine import FlowchartLayoutManager

# Vertical layout
elements = [
    {"id": "start", "flowchart_type": "start", "text": "Start"},
    {"id": "process", "flowchart_type": "process", "text": "Process"},
    {"id": "end", "flowchart_type": "end", "text": "End"}
]

positioned = FlowchartLayoutManager.create_vertical_layout(
    elements, start_x=2, start_y=1, spacing_y=1.5
)

# Auto-connect sequential elements
connections = FlowchartLayoutManager.auto_connect_sequential(
    ["start", "process", "end"]
)
```

## Examples

See the complete examples in:
- `examples/flowchart_example.json` - Comprehensive flowchart demonstrations
- `examples/test_flowchart.py` - Python test script

## Testing

Run the test script to verify flowchart functionality:

```bash
cd src/examples
python test_flowchart.py
```

This will generate sample flowchart presentations in the `output/` directory.

## Integration

The flowchart functionality integrates seamlessly with existing pypptx-engine features:

- Use alongside text, images, charts, and tables
- Apply the same formatting options (fonts, colors, shadows)
- Include in multi-slide presentations
- Combine with existing shape types

## Advanced Features

### Complex Business Processes

Create sophisticated workflows with multiple decision points, parallel processes, and various shape types.

### Custom Connection Routing

Fine-tune connection paths with specific attachment points and connector types.

### Dynamic Layouts

Use layout utilities to programmatically generate flowcharts based on data structures.

### Professional Styling

Apply consistent themes across entire presentations with customizable color schemes and typography.
