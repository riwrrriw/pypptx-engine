# pypptx-engine Architecture

## System Architecture

pypptx-engine follows a layered architecture with clear separation of concerns:

```
┌─────────────────────────────────────────────────────────────┐
│                        CLI Layer                            │
│                     (cli.py)                               │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                    Engine Layer                             │
│                   (engine.py)                              │
│  • PPTXEngine - Main orchestrator                          │
│  • Presentation properties management                       │
│  • Slide size configuration                                │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                   Manager Layer                             │
│                  (slides.py)                               │
│  • SlideManager - Slide creation & layout                  │
│  • Background application                                   │
│  • Placeholder management                                   │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                   Factory Layer                             │
│                  (shapes.py)                               │
│  • ShapeFactory - Shape creation coordination               │
│  • TextShapeHandler - Text & bullet shapes                 │
│  • ImageShapeHandler - Image shapes                         │
│  • ChartShapeHandler - Chart shapes                         │
│  • TableShapeHandler - Table shapes                         │
│  • AutoShapeHandler - Auto shapes & connectors             │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                 Formatting Layer                            │
│                (formatters.py)                             │
│  • ColorFormatter - Color parsing & application            │
│  • FontFormatter - Font & text formatting                  │
│  • LineFormatter - Border & line formatting                │
│  • ShadowFormatter - Shadow effects                        │
└─────────────────────────────────────────────────────────────┘
```

## Core Components

### PPTXEngine (`engine.py`)
- **Purpose**: Main orchestrator for the entire conversion process
- **Responsibilities**:
  - Initialize all sub-components
  - Parse JSON configuration
  - Apply presentation-level properties
  - Coordinate slide creation
- **Key Methods**:
  - `create_presentation()`: Main entry point
  - `_apply_presentation_properties()`: Set metadata
  - `_apply_slide_size()`: Configure dimensions

### SlideManager (`slides.py`)
- **Purpose**: Handle slide-level operations
- **Responsibilities**:
  - Create slides with specified layouts
  - Apply slide backgrounds (solid, gradient, picture)
  - Manage placeholder content
  - Coordinate shape placement
- **Key Methods**:
  - `create_slide()`: Create and configure slides
  - `_apply_background()`: Set slide backgrounds
  - `_fill_placeholders()`: Populate slide placeholders

### ShapeFactory (`shapes.py`)
- **Purpose**: Central factory for creating all shape types
- **Responsibilities**:
  - Route shape creation to appropriate handlers
  - Manage position and size calculations
  - Coordinate formatting application
- **Shape Handlers**:
  - `TextShapeHandler`: Text boxes and bullet lists
  - `ImageShapeHandler`: Picture insertion
  - `ChartShapeHandler`: Data visualization
  - `TableShapeHandler`: Structured data
  - `AutoShapeHandler`: Geometric shapes and connectors

### Formatters (`formatters.py`)
- **Purpose**: Apply visual formatting to shapes
- **Components**:
  - `ColorFormatter`: Parse colors (hex, RGB) and apply fills
  - `FontFormatter`: Font properties and paragraph formatting
  - `LineFormatter`: Borders and line styles
  - `ShadowFormatter`: Shadow effects

## Data Flow

1. **Input**: JSON configuration file
2. **Parsing**: CLI loads and validates JSON
3. **Engine**: PPTXEngine creates presentation object
4. **Slides**: SlideManager creates each slide
5. **Shapes**: ShapeFactory creates shapes on slides
6. **Formatting**: Formatters apply visual properties
7. **Output**: PPTX file saved to disk

## Design Patterns

### Factory Pattern
- `ShapeFactory` creates appropriate shape handlers
- Each handler specializes in one shape type
- Extensible for new shape types

### Strategy Pattern
- Different formatting strategies for colors, fonts, lines
- Pluggable formatters for different visual properties

### Builder Pattern
- Step-by-step construction of presentations
- Configurable at each level (presentation, slide, shape)

## Extension Points

### Adding New Shape Types
1. Create new handler class in `shapes.py`
2. Implement shape creation logic
3. Register in `ShapeFactory.create_shape()`
4. Add formatting support if needed

### Adding New Formatters
1. Create formatter class in `formatters.py`
2. Implement static formatting methods
3. Integrate with shape handlers

### Custom Layouts
1. Extend `SlideManager` with new layout logic
2. Add layout-specific configuration options
3. Implement in `create_slide()` method

## Dependencies

- **python-pptx**: Core PowerPoint manipulation library
- **Standard Library**: JSON parsing, file operations, argument parsing
- **No External Dependencies**: Minimal dependency footprint for easy deployment