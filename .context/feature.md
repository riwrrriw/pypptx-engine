# pypptx-engine Features

## Comprehensive Shape Support

### Text Shapes
- **Rich Text Boxes**: Full formatting control with fonts, colors, and alignment
- **Multi-paragraph Support**: Each paragraph can have individual formatting
- **Text Frame Properties**: Margins, word wrap, vertical anchoring
- **Advanced Typography**: Bold, italic, underline, custom font families

### Bullet Lists
- **Multi-level Lists**: Support for nested bullet points
- **Custom Styling**: Font formatting for each bullet level
- **Flexible Content**: String or array-based item definitions

### Images
- **Multiple Formats**: Support for common image formats (PNG, JPG, etc.)
- **Path Resolution**: Automatic handling of relative and absolute paths
- **Sizing Options**: Explicit width/height or automatic scaling
- **Asset Management**: Base directory configuration for organized assets

### Charts
- **Chart Types**: Column (clustered), Pie, and extensible for more types
- **Data Series**: Multiple data series with custom names and values
- **Formatting**: Titles, legends, and axis configuration
- **Categories**: Custom category labels for data visualization

### Tables
- **Dynamic Sizing**: Configurable rows and columns
- **Cell Formatting**: Individual cell styling and content
- **Data Population**: Array-based data structure for easy population
- **Table Styles**: Built-in table formatting options

### Auto Shapes
- **Geometric Shapes**: Rectangle, oval, diamond, rounded rectangle
- **Text Integration**: Add text content to shapes with formatting
- **Custom Styling**: Fill colors, borders, and shadow effects

### Connectors
- **Line Types**: Straight connectors with customizable styling
- **Positioning**: Flexible start and end point configuration
- **Line Formatting**: Color, width, and style options

## Advanced Formatting

### Color System
- **Hex Colors**: Standard hex color notation (#RRGGBB)
- **RGB Colors**: Direct RGB value specification
- **Color Parsing**: Flexible color input formats
- **Theme Integration**: Support for presentation color themes

### Fill Options
- **Solid Fills**: Single color backgrounds
- **Gradient Fills**: Multi-stop gradient effects
- **Pattern Fills**: Textured background options
- **Picture Fills**: Image-based backgrounds
- **Transparency**: No-fill options for transparent shapes

### Font Formatting
- **Font Families**: Custom font selection
- **Font Sizes**: Point-based sizing
- **Font Styles**: Bold, italic, underline combinations
- **Font Colors**: Full color specification support

### Paragraph Formatting
- **Text Alignment**: Left, center, right, justify
- **Line Spacing**: Custom line height control
- **Spacing**: Before and after paragraph spacing
- **Indentation**: Paragraph-level indentation control

### Visual Effects
- **Shadows**: Configurable shadow effects with color and visibility
- **Borders**: Line width, color, and style options
- **Transparency**: Alpha channel support for various elements

## Presentation Configuration

### Presentation Properties
- **Metadata**: Title, author, subject, keywords, comments
- **Document Properties**: Category and other document attributes
- **Core Properties**: Integration with PowerPoint's built-in properties

### Slide Management
- **Layout Selection**: Choose from PowerPoint's built-in layouts
- **Custom Sizing**: Inch or centimeter-based dimensions
- **Background Options**: Solid colors, gradients, or images
- **Placeholder Support**: Automatic placeholder population

### Asset Management
- **Base Directory**: Configurable asset root for relative paths
- **Path Resolution**: Automatic resolution of image and asset paths
- **File Validation**: Existence checking for referenced assets

## CLI Features

### Command Line Interface
- **Input/Output**: Flexible file path specification
- **Asset Base**: Configurable base directory for assets
- **Error Handling**: Comprehensive error reporting and validation
- **Progress Feedback**: Clear success/failure indicators

### Usage Examples
```bash
# Basic usage
pypptx-engine --input presentation.json --output result.pptx

# With custom asset directory
pypptx-engine -i config.json -o slides.pptx --assets-base ./images/

# Using short flags
pypptx-engine -i input.json -o output.pptx
```

## JSON Configuration Schema

### Presentation Structure
```json
{
  "presentation": {
    "properties": {
      "title": "Presentation Title",
      "author": "Author Name",
      "subject": "Subject",
      "keywords": "keyword1, keyword2"
    },
    "size": {
      "width_in": 16,
      "height_in": 9
    },
    "slides": [...]
  }
}
```

### Slide Configuration
```json
{
  "layout": 6,
  "background": "#ffffff",
  "shapes": [...],
  "placeholders": {...}
}
```

### Shape Examples
```json
{
  "type": "text",
  "text": "Sample Text",
  "x": 1, "y": 2, "w": 8, "h": 1,
  "font": {
    "name": "Arial",
    "size": 24,
    "bold": true,
    "color": "#333333"
  },
  "fill": {
    "type": "gradient",
    "stops": [
      {"color": "#ff0000", "position": 0},
      {"color": "#0000ff", "position": 1}
    ]
  }
}
```

## Integration Capabilities

### Programmatic Usage
- **Python API**: Direct integration into Python applications
- **Engine Class**: `PPTXEngine` for programmatic presentation creation
- **Configuration Objects**: Dictionary-based configuration for dynamic generation

### Workflow Integration
- **CI/CD Pipelines**: Automated presentation generation
- **Data Processing**: Convert data analysis results to presentations
- **Template Systems**: Dynamic presentation generation from templates
- **Batch Processing**: Generate multiple presentations from data sets

## Extensibility

### Custom Shape Types
- **Handler Pattern**: Easy addition of new shape types
- **Formatting Integration**: Automatic formatting support for new shapes
- **Configuration Schema**: Extensible JSON schema for new features

### Custom Formatters
- **Pluggable Architecture**: Add new formatting capabilities
- **Consistent API**: Uniform formatting interface across shape types
- **Reusable Components**: Share formatting logic across different shapes