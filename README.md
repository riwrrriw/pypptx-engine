# pypptx-engine

A powerful Python library for converting JSON specifications into professional PowerPoint presentations using python-pptx.

## Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd pypptx-engine

# Install dependencies using Poetry
poetry install

# Or install using pip
pip install -e .
```

### Basic Usage

```bash
# Generate a presentation from JSON
pypptx-engine --input src/examples/comprehensive_example.json --output src/output/demo.pptx

# With custom asset directory
pypptx-engine -i src/examples/comprehensive_example.json -o demo.pptx --assets-base ./assets/
```

### Python API

```python
from pypptx_engine import PPTXEngine
import json

# Load configuration
with open('presentation.json', 'r') as f:
    config = json.load(f)

# Create presentation
engine = PPTXEngine()
presentation = engine.create_presentation(config)

# Save to file
presentation.save('output.pptx')
```

## Features

- **JSON-Driven**: Define entire presentations using structured JSON
- **Comprehensive Shape Support**: Text, bullets, images, charts, tables, auto shapes
- **Advanced Formatting**: Fonts, colors, gradients, shadows, borders
- **CLI Tool**: Command-line interface for easy integration
- **Asset Management**: Automatic path resolution for images and resources
- **Extensible Architecture**: Easy to add new shape types and formatters

## Supported Shape Types

| Shape Type | Description | Features |
|------------|-------------|----------|
| **Text** | Rich text boxes | Multi-paragraph, full formatting |
| **Bullet** | Bullet lists | Multi-level, custom styling |
| **Image** | Pictures | Multiple formats, auto-sizing |
| **Chart** | Data visualization | Column, pie charts with legends |
| **Table** | Structured data | Dynamic sizing, cell formatting |
| **AutoShape** | Geometric shapes | Rectangle, oval, diamond, etc. |
| **Connector** | Lines and arrows | Customizable styling |

## JSON Configuration

### Basic Structure

```json
{
  "presentation": {
    "properties": {
      "title": "My Presentation",
      "author": "Author Name"
    },
    "size": { "width_in": 16, "height_in": 9 },
    "slides": [
      {
        "layout": 6,
        "background": "#ffffff",
        "shapes": [
          {
            "type": "text",
            "text": "Hello World",
            "x": 2, "y": 2, "w": 8, "h": 1,
            "font": {
              "name": "Arial",
              "size": 24,
              "bold": true,
              "color": "#333333"
            }
          }
        ]
      }
    ]
  }
}
```

### Advanced Features

```json
{
  "type": "text",
  "text": "Styled Text",
  "x": 1, "y": 2, "w": 8, "h": 2,
  "font": {
    "name": "Calibri",
    "size": 20,
    "bold": true,
    "color": "#2c3e50"
  },
  "fill": {
    "type": "gradient",
    "stops": [
      {"color": "#3498db", "position": 0},
      {"color": "#2980b9", "position": 1}
    ]
  },
  "shadow": {
    "visible": true,
    "color": "#000000"
  },
  "line": {
    "color": "#34495e",
    "width": 2
  }
}
```

## Examples

See `src/examples/comprehensive_example.json` for a complete demonstration of all features including:

- Title slides with gradient backgrounds
- Text formatting and bullet lists
- Charts with multiple data series
- Tables with custom styling
- Auto shapes with various fills
- Advanced formatting effects

## Architecture

pypptx-engine uses a modular architecture:

- **PPTXEngine**: Main orchestrator
- **SlideManager**: Slide creation and layout
- **ShapeFactory**: Shape creation coordination
- **Formatters**: Visual formatting (colors, fonts, effects)

For detailed architecture information, see `.context/architecture.md`.

## Development

### Project Structure

```
pypptx-engine/
├── src/
│   ├── pypptx_engine/          # Main package
│   │   ├── engine.py           # Core engine
│   │   ├── slides.py           # Slide management
│   │   ├── shapes.py           # Shape creation
│   │   ├── formatters.py       # Visual formatting
│   │   └── cli.py              # Command line interface
│   ├── examples/               # Example configurations
│   └── output/                 # Generated presentations
├── tests/                      # Test suite
├── .context/                   # Documentation
│   ├── overview.md            # Project overview
│   ├── architecture.md        # Architecture details
│   └── feature.md             # Feature documentation
└── README.md                  # This file
```

### Running Tests

```bash
# Run all tests
poetry run pytest

# Run with coverage
poetry run pytest --cov=pypptx_engine
```

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite
6. Submit a pull request

## Requirements

- Python 3.13+
- python-pptx >= 1.0.2

## License

[Add your license information here]

## Support

For questions, issues, or contributions:
- Create an issue on GitHub
- Check the documentation in `.context/`
- Review example configurations in `src/examples/`


https://learn.microsoft.com/en-us/office/open-xml/presentation/overview