"""
Shape creation and management for all supported python-pptx shape types
"""
from __future__ import annotations

import os
from typing import Any, Dict, List

from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

from .formatters import FontFormatter, ColorFormatter, LineFormatter, ShadowFormatter


class ShapeFactory:
    """Factory for creating various types of shapes."""
    
    def __init__(self, color_formatter):
        self.color_formatter = color_formatter
        self.text_handler = TextShapeHandler(color_formatter)
        self.image_handler = ImageShapeHandler()
        self.chart_handler = ChartShapeHandler()
        self.table_handler = TableShapeHandler(color_formatter)
        self.autoshape_handler = AutoShapeHandler(color_formatter)
    
    def create_shape(self, slide, shape_config: Dict[str, Any], base_dir: str) -> None:
        """Create a shape based on configuration."""
        shape_type = shape_config.get("type", "").lower()
        
        # Get position and size
        x = Inches(shape_config.get("x", 0))
        y = Inches(shape_config.get("y", 0))
        w = Inches(shape_config.get("w", 4))
        h = Inches(shape_config.get("h", 1))
        
        if shape_type == "text":
            self.text_handler.create_text_shape(slide, shape_config, x, y, w, h)
        elif shape_type == "bullet":
            self.text_handler.create_bullet_shape(slide, shape_config, x, y, w, h)
        elif shape_type == "image":
            self.image_handler.create_image_shape(slide, shape_config, x, y, w, h, base_dir)
        elif shape_type == "chart":
            self.chart_handler.create_chart_shape(slide, shape_config, x, y, w, h)
        elif shape_type == "table":
            self.table_handler.create_table_shape(slide, shape_config, x, y, w, h)
        elif shape_type == "autoshape":
            self.autoshape_handler.create_autoshape(slide, shape_config, x, y, w, h)
        elif shape_type == "connector":
            self.autoshape_handler.create_connector(slide, shape_config, x, y, w, h)


class TextShapeHandler:
    """Handle text-based shapes including textboxes and bullet lists."""
    
    def __init__(self, color_formatter):
        self.color_formatter = color_formatter
    
    def create_text_shape(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create a text box shape."""
        textbox = slide.shapes.add_textbox(x, y, w, h)
        text_frame = textbox.text_frame
        
        # Clear default content
        text_frame.clear()
        
        # Set text frame properties
        self._apply_text_frame_formatting(text_frame, config.get("text_frame", {}))
        
        # Add text content
        text_content = config.get("text", "")
        if isinstance(text_content, str):
            # Simple text
            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = text_content
            
            # Apply formatting
            FontFormatter.apply_font_formatting(run.font, config.get("font", {}))
            FontFormatter.apply_paragraph_formatting(p, config.get("paragraph", {}))
        elif isinstance(text_content, list):
            # Multiple paragraphs
            for i, para_text in enumerate(text_content):
                p = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                run = p.add_run()
                run.text = str(para_text)
                
                # Apply formatting
                FontFormatter.apply_font_formatting(run.font, config.get("font", {}))
                FontFormatter.apply_paragraph_formatting(p, config.get("paragraph", {}))
        
        # Apply shape-level formatting
        self._apply_shape_formatting(textbox, config)
    
    def create_bullet_shape(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create a bullet list shape."""
        textbox = slide.shapes.add_textbox(x, y, w, h)
        text_frame = textbox.text_frame
        text_frame.clear()
        
        # Set text frame properties
        self._apply_text_frame_formatting(text_frame, config.get("text_frame", {}))
        
        items = config.get("items", [])
        for i, item in enumerate(items):
            p = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
            p.text = str(item)
            p.level = config.get("level", 0)  # Bullet level
            
            # Apply formatting
            FontFormatter.apply_font_formatting(p.font, config.get("font", {}))
            FontFormatter.apply_paragraph_formatting(p, config.get("paragraph", {}))
        
        # Apply shape-level formatting
        self._apply_shape_formatting(textbox, config)
    
    def _apply_text_frame_formatting(self, text_frame, config: Dict[str, Any]) -> None:
        """Apply text frame specific formatting."""
        if not config:
            return
        
        if "margin_left" in config:
            text_frame.margin_left = Inches(config["margin_left"])
        if "margin_right" in config:
            text_frame.margin_right = Inches(config["margin_right"])
        if "margin_top" in config:
            text_frame.margin_top = Inches(config["margin_top"])
        if "margin_bottom" in config:
            text_frame.margin_bottom = Inches(config["margin_bottom"])
        
        if "word_wrap" in config:
            text_frame.word_wrap = bool(config["word_wrap"])
        
        if "auto_size" in config:
            # Handle auto-sizing options
            pass
        
        vertical_anchor = config.get("vertical_anchor", "").lower()
        if vertical_anchor == "top":
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        elif vertical_anchor == "middle":
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        elif vertical_anchor == "bottom":
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.BOTTOM
    
    def _apply_shape_formatting(self, shape, config: Dict[str, Any]) -> None:
        """Apply general shape formatting."""
        # Fill formatting
        if "fill" in config:
            ColorFormatter.apply_fill(shape, config["fill"])
        
        # Line formatting
        if "line" in config:
            LineFormatter.apply_line_formatting(shape.line, config["line"])
        
        # Shadow formatting
        if "shadow" in config:
            ShadowFormatter.apply_shadow(shape, config["shadow"])


class ImageShapeHandler:
    """Handle image shapes."""
    
    def create_image_shape(self, slide, config: Dict[str, Any], x, y, w, h, base_dir: str) -> None:
        """Create an image shape."""
        image_path = config.get("path")
        if not image_path:
            print("[WARN] No image path specified")
            return
        
        # Resolve path
        if not os.path.isabs(image_path):
            image_path = os.path.join(base_dir, image_path)
        
        if not os.path.exists(image_path):
            print(f"[WARN] Image not found, skipping: {image_path}")
            return
        
        # Determine dimensions
        width = Inches(config["w"]) if "w" in config else None
        height = Inches(config["h"]) if "h" in config else None
        
        # Add picture
        picture = slide.shapes.add_picture(image_path, x, y, width=width, height=height)
        
        # Apply formatting
        if "line" in config:
            LineFormatter.apply_line_formatting(picture.line, config["line"])
        
        if "shadow" in config:
            ShadowFormatter.apply_shadow(picture, config["shadow"])


class ChartShapeHandler:
    """Handle chart shapes."""
    
    def create_chart_shape(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create a chart shape."""
        chart_type_name = config.get("chartType", "COLUMN_CLUSTERED")
        
        try:
            chart_type = getattr(XL_CHART_TYPE, chart_type_name)
        except AttributeError:
            print(f"[WARN] Unsupported chart type: {chart_type_name}")
            return
        
        # Prepare chart data
        chart_data = CategoryChartData()
        
        categories = config.get("categories", [])
        chart_data.categories = categories
        
        for series_config in config.get("series", []):
            name = series_config.get("name", "Series")
            values = series_config.get("values", [])
            chart_data.add_series(name, values)
        
        # Add chart to slide
        chart_shape = slide.shapes.add_chart(chart_type, x, y, w, h, chart_data)
        chart = chart_shape.chart
        
        # Apply chart formatting
        self._apply_chart_formatting(chart, config.get("formatting", {}))
    
    def _apply_chart_formatting(self, chart, formatting: Dict[str, Any]) -> None:
        """Apply chart-specific formatting."""
        if not formatting:
            return
        
        # Chart title
        if "title" in formatting:
            chart.has_title = True
            chart.chart_title.text_frame.text = formatting["title"]
        
        # Legend
        if "legend" in formatting:
            legend_config = formatting["legend"]
            chart.has_legend = legend_config.get("visible", True)
            if chart.has_legend:
                legend = chart.legend
                position = legend_config.get("position", "right").upper()
                # Set legend position based on config
        
        # Axes formatting
        if "axes" in formatting:
            axes_config = formatting["axes"]
            if "category" in axes_config:
                # Format category axis
                pass
            if "value" in axes_config:
                # Format value axis
                pass


class TableShapeHandler:
    """Handle table shapes."""
    
    def __init__(self, color_formatter):
        self.color_formatter = color_formatter
    
    def create_table_shape(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create a table shape."""
        rows = config.get("rows", 2)
        cols = config.get("cols", 2)
        
        # Add table
        table_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
        table = table_shape.table
        
        # Fill table data
        data = config.get("data", [])
        for row_idx, row_data in enumerate(data):
            if row_idx >= rows:
                break
            for col_idx, cell_data in enumerate(row_data):
                if col_idx >= cols:
                    break
                
                cell = table.cell(row_idx, col_idx)
                if isinstance(cell_data, str):
                    cell.text = cell_data
                elif isinstance(cell_data, dict):
                    cell.text = cell_data.get("text", "")
                    # Apply cell formatting
                    self._apply_cell_formatting(cell, cell_data.get("formatting", {}))
        
        # Apply table-level formatting
        self._apply_table_formatting(table, config.get("formatting", {}))
    
    def _apply_cell_formatting(self, cell, formatting: Dict[str, Any]) -> None:
        """Apply formatting to a table cell."""
        if not formatting:
            return
        
        # Fill formatting
        if "fill" in formatting:
            ColorFormatter.apply_fill(cell, formatting["fill"])
        
        # Text formatting
        if "font" in formatting:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    FontFormatter.apply_font_formatting(run.font, formatting["font"])
    
    def _apply_table_formatting(self, table, formatting: Dict[str, Any]) -> None:
        """Apply table-level formatting."""
        if not formatting:
            return
        
        # Table style
        if "style" in formatting:
            # Apply table style
            pass


class AutoShapeHandler:
    """Handle auto shapes and connectors."""
    
    def __init__(self, color_formatter):
        self.color_formatter = color_formatter
    
    def create_autoshape(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create an auto shape."""
        shape_type_name = config.get("shape_type", "RECTANGLE")
        
        try:
            shape_type = getattr(MSO_SHAPE, shape_type_name)
        except AttributeError:
            print(f"[WARN] Unsupported auto shape type: {shape_type_name}")
            return
        
        # Add auto shape
        autoshape = slide.shapes.add_shape(shape_type, x, y, w, h)
        
        # Add text if specified
        text = config.get("text")
        if text and autoshape.has_text_frame:
            autoshape.text = text
            
            # Apply text formatting
            if "font" in config:
                for paragraph in autoshape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        FontFormatter.apply_font_formatting(run.font, config["font"])
        
        # Apply shape formatting
        if "fill" in config:
            ColorFormatter.apply_fill(autoshape, config["fill"])
        
        if "line" in config:
            LineFormatter.apply_line_formatting(autoshape.line, config["line"])
        
        if "shadow" in config:
            ShadowFormatter.apply_shadow(autoshape, config["shadow"])
    
    def create_connector(self, slide, config: Dict[str, Any], x, y, w, h) -> None:
        """Create a connector shape."""
        connector_type = config.get("connector_type", "STRAIGHT")
        
        # Add connector
        begin_x, begin_y = x, y
        end_x, end_y = x + w, y + h
        
        # Note: python-pptx connector API is limited
        # This is a simplified implementation
        connector = slide.shapes.add_connector(
            1, begin_x, begin_y, end_x, end_y  # STRAIGHT connector type
        )
        
        # Apply line formatting
        if "line" in config:
            LineFormatter.apply_line_formatting(connector.line, config["line"])
