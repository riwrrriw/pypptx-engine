# AI/RAG System for pypptx-engine

An intelligent content generation system that transforms documents into beautiful PowerPoint presentations through a 3-step AI-powered process.

## 🚀 Quick Start with Poetry

```bash
# Navigate to the AI directory
cd src/ai

# Install Poetry (if not already installed)
curl -sSL https://install.python-poetry.org | python3 -

# Install dependencies with Poetry
poetry install

# For additional features (data processing)
poetry install --extras data

# Set up environment variables
cp .env.example .env
# Edit .env with your Azure OpenAI API key

# Run the full 3-step process
poetry run python cli.py my_presentation --all

# Or run step by step
poetry run python cli.py my_presentation --step 1  # Generate content plan
poetry run python cli.py my_presentation --step 2  # Generate JSON
poetry run python cli.py my_presentation --step 3  # Generate PPTX
```

## 📦 Installation Options

### Option 1: Poetry (Recommended)
```bash
# Install core dependencies (includes Azure OpenAI)
poetry install

# Or install with data processing features
poetry install --extras data      # Adds pandas, numpy
poetry install --extras full      # All optional features
```

### Option 2: pip (Alternative)
```bash
# Install basic dependencies
pip install -r requirements.txt

# Core dependencies include Azure OpenAI SDK and python-pptx
```

## 📁 Directory Structure

```
ai/
├── cli.py                 # Main CLI interface
├── content_extractor.py   # Document content extraction
├── json_generator.py      # JSON specification generation
├── ai_foundry.py         # AI model integration
├── content_plan.md       # Content plan template
├── resource/             # Place your source documents here
│   ├── .gitkeep
│   └── [your PDFs, DOCX, TXT, MD files]
└── requirements.txt      # Python dependencies
```

## 🔄 3-Step Process

### Step 1: Content Analysis & Planning
- Extracts content from documents in `resource/` folder
- Supports: PDF, DOCX, TXT, MD, JSON files
- Generates structured `content_plan.md`
- **User can review and customize before proceeding**

### Step 2: JSON Generation
- Converts content plan to pypptx-engine JSON specification
- Applies professional design patterns
- Supports multiple slide types: text, bullets, images, charts, tables, flowcharts
- Generates beautiful, consistent styling

### Step 3: PPTX Creation
- Uses pypptx-engine to create final PowerPoint presentation
- Saves to `result/<work_name>/` directory
- Includes project summary and resource mapping

## 🤖 Azure OpenAI Integration

The system is powered by Azure OpenAI GPT-5 for intelligent content generation:

### AI Model
- **Azure OpenAI GPT-5**: Advanced language model for professional presentation content
- **Deployment**: Custom Azure deployment (`gpt-5-for-pptx`)
- **API Version**: `2024-12-01-preview` for latest features

### Setup Azure OpenAI
```bash
# Set your Azure OpenAI credentials
export AI_FOUNDRY_API_KEY="your-azure-openai-api-key"
export AI_FOUNDRY_BASE_URL="https://your-resource.cognitiveservices.azure.com/"
export AI_MODEL_DEPLOYMENT="gpt-5-for-pptx"

# Or create .env file from template
cp .env.example .env
# Edit .env with your Azure credentials
```

### AI Features
- **Content Enhancement**: Improves extracted content summaries
- **Design Suggestions**: AI-powered color schemes and layouts
- **Slide Optimization**: Better titles and bullet points
- **Visual Recommendations**: Suggests appropriate charts and images

## 📋 Usage Examples

### Basic Usage
```bash
# Place documents in resource/ folder
cp ~/Documents/my_report.pdf resource/
cp ~/Documents/data.xlsx resource/

# Generate presentation
python cli.py quarterly_report --all
```

### Advanced Usage
```bash
# Custom resource directory
python cli.py project_name --step 1 --resource-dir /path/to/docs

# Use existing content plan
python cli.py project_name --step 2 --content-plan /path/to/plan.md

# Use existing JSON spec
python cli.py project_name --step 3 --json-spec /path/to/spec.json
```

## 📄 Supported File Formats

| Format | Extension | Features |
|--------|-----------|----------|
| PDF | `.pdf` | Text extraction, page count |
| Word | `.docx`, `.doc` | Paragraph extraction, formatting |
| Text | `.txt` | Plain text processing |
| Markdown | `.md` | Header structure, formatting |
| JSON | `.json` | Data structure analysis |

## 🎨 Output Structure

```
result/
└── <work_name>/
    ├── content_plan.md           # Generated content plan
    ├── <work_name>_presentation.json  # JSON specification
    ├── <work_name>.pptx         # Final PowerPoint file
    └── project_summary.md       # Project overview and mapping
```

## ⚙️ Configuration

### Content Plan Customization
Edit `content_plan.md` after step 1 to:
- Modify slide titles and content
- Adjust design preferences
- Change color schemes
- Add specific visual requirements

### Design Themes
- **Professional**: Corporate blue and white
- **Modern**: Contemporary colors with gradients
- **Creative**: Vibrant colors with dynamic layouts

## 🔧 Dependencies & Environment Setup

### Poetry Configuration
The project uses Poetry for dependency management with optional extras:

- **Base**: Core functionality with Azure OpenAI and python-pptx
- **data**: Data processing libraries (pandas, numpy)
- **full**: All optional features combined

**Recent Fix**: Added `python-pptx` dependency to resolve PPTX generation issues.

### Environment Variables
Copy `.env.example` to `.env` and configure:

```bash
# Required for AI features
AI_FOUNDRY_API_KEY=your-azure-openai-api-key-here
AI_FOUNDRY_BASE_URL=https://ai-peerawatr-6647.cognitiveservices.azure.com/
AI_FOUNDRY_API_VERSION=2024-12-01-preview
AI_MODEL_DEPLOYMENT=gpt-5-for-pptx
AI_MODEL=gpt-5

# System configuration
DEFAULT_THEME=professional
ENABLE_AI_ENHANCEMENT=true
```

### Azure OpenAI Setup
1. **Azure OpenAI**: Get your API key from Azure OpenAI resource
2. **Deployment**: Use your deployment name (e.g., `gpt-5-for-pptx`)
3. **Endpoint**: Your Azure cognitive services endpoint

## 🚨 Troubleshooting

### Common Issues

**No resources found**
```bash
# Ensure files are in resource/ directory
ls resource/
# Add supported file types (.pdf, .docx, .txt, .md, .json)
```

**AI enhancement not working**
```bash
# Check Azure OpenAI API key
echo $AI_FOUNDRY_API_KEY
# Install dependencies
poetry install
```

**PPTX generation fails**
```bash
# Ensure python-pptx is installed in AI environment
poetry install  # This now includes python-pptx dependency

# Verify pypptx-engine module is accessible
cd ../../
python -m pypptx_engine.cli --help
```

## 🎯 Best Practices

1. **Resource Organization**: Place related documents in `resource/` folder
2. **Content Review**: Always review and customize `content_plan.md` in step 1
3. **Iterative Improvement**: Re-run steps 2-3 after content plan changes
4. **AI Enhancement**: Use AI features for professional content improvement
5. **Output Management**: Check `result/` directory for all generated files

## 🔮 Future Enhancements

- [ ] Real-time collaboration features
- [ ] Template library integration
- [ ] Advanced chart generation from data
- [ ] Multi-language support
- [ ] Custom animation sequences
- [ ] Brand guideline integration

---

**Generated by pypptx-engine AI system** | [Documentation](../docs/) | [Examples](../examples/)
