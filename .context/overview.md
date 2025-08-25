# pypptx-engine Overview

## What is pypptx-engine?

pypptx-engine is a powerful Python library that converts JSON specifications into professional PowerPoint presentations. Built on top of the `python-pptx` library, it provides a declarative approach to creating presentations programmatically.

## Key Features

- **JSON-Driven**: Define entire presentations using structured JSON configuration
- **CLI Tool**: Command-line interface for easy integration into workflows
- **Comprehensive Shape Support**: Text boxes, bullet lists, images, charts, tables, and auto shapes
- **Advanced Formatting**: Fonts, colors, gradients, shadows, borders, and more
- **Modular Architecture**: Clean separation of concerns with dedicated handlers for each shape type
- **Asset Management**: Automatic resolution of relative paths for images and other assets

## Use Cases

- **Automated Reporting**: Generate presentations from data automatically
- **Template-Based Presentations**: Create consistent presentations from JSON templates
- **Batch Processing**: Generate multiple presentations with different data
- **Integration**: Embed presentation generation into larger applications
- **Prototyping**: Quickly create presentation mockups from configuration files

## Core Components

1. **PPTXEngine**: Main orchestrator that coordinates the conversion process
2. **SlideManager**: Handles slide creation, layouts, and backgrounds
3. **ShapeFactory**: Creates different types of shapes based on configuration
4. **Formatters**: Apply visual formatting (colors, fonts, lines, shadows)
5. **CLI**: Command-line interface for standalone usage

## Supported Shape Types

- **Text**: Rich text boxes with full formatting control
- **Bullet Lists**: Multi-level bullet points with custom styling
- **Images**: Picture insertion with sizing and positioning
- **Charts**: Column, pie, and other chart types with data visualization
- **Tables**: Structured data presentation with cell formatting
- **Auto Shapes**: Rectangles, circles, diamonds, and other geometric shapes
- **Connectors**: Lines and arrows for connecting elements

## Architecture Philosophy

The project follows a modular design where each component has a single responsibility:
- **Separation of Concerns**: Each shape type has its own handler
- **Extensibility**: Easy to add new shape types or formatting options
- **Maintainability**: Clear code organization with focused modules
- **Reusability**: Components can be used independently or together