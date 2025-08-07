# Installation Guide

This guide provides detailed instructions for installing Offitrans and its dependencies.

## Quick Installation

### Install from PyPI (Coming Soon)

```bash
pip install offitrans
```

### Install from Source

```bash
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans
pip install -e .
```

## Dependencies

Offitrans has both required and optional dependencies depending on which features you want to use.

### Required Dependencies

These are automatically installed with Offitrans:

```bash
pip install requests>=2.28.0
```

### Optional Dependencies

Install these based on the file formats you want to process:

#### Excel Support
```bash
pip install openpyxl>=3.0.10
```

#### Word Document Support
```bash
pip install python-docx>=0.8.11
```

#### PDF Support
```bash
pip install PyPDF2>=3.0.1
```

#### PowerPoint Support
```bash
pip install python-pptx>=0.6.21
```

#### Image Processing (Recommended)
```bash
pip install Pillow>=9.0.0
```

#### Enhanced Text Processing
```bash
pip install lxml>=4.9.0
```

### Development Dependencies

If you plan to contribute to Offitrans:

```bash
pip install -r requirements-dev.txt
```

Or install individual development tools:

```bash
pip install pytest>=7.0.0 pytest-cov>=4.0.0 black>=22.0.0 flake8>=5.0.0
```

## Platform-Specific Instructions

### Windows

1. **Install Python 3.7+** from [python.org](https://www.python.org/)

2. **Install Offitrans**:
   ```cmd
   pip install offitrans
   ```

3. **For Excel support with images**:
   ```cmd
   pip install openpyxl Pillow
   ```

4. **Common issues**:
   - If you get SSL errors, try: `pip install --trusted-host pypi.org --trusted-host pypi.python.org offitrans`
   - For corporate networks, you may need to configure proxy settings

### macOS

1. **Install Python 3.7+** using Homebrew:
   ```bash
   brew install python
   ```

2. **Install Offitrans**:
   ```bash
   pip3 install offitrans
   ```

3. **For full functionality**:
   ```bash
   pip3 install openpyxl python-docx PyPDF2 python-pptx Pillow
   ```

### Linux (Ubuntu/Debian)

1. **Install Python and pip**:
   ```bash
   sudo apt update
   sudo apt install python3 python3-pip
   ```

2. **Install system dependencies**:
   ```bash
   sudo apt install python3-dev libxml2-dev libxslt1-dev zlib1g-dev
   ```

3. **Install Offitrans**:
   ```bash
   pip3 install offitrans
   ```

### Linux (CentOS/RHEL/Fedora)

1. **Install Python and pip**:
   ```bash
   # CentOS/RHEL
   sudo yum install python3 python3-pip
   
   # Fedora
   sudo dnf install python3 python3-pip
   ```

2. **Install system dependencies**:
   ```bash
   # CentOS/RHEL
   sudo yum install python3-devel libxml2-devel libxslt-devel zlib-devel
   
   # Fedora
   sudo dnf install python3-devel libxml2-devel libxslt-devel zlib-devel
   ```

3. **Install Offitrans**:
   ```bash
   pip3 install offitrans
   ```

## Virtual Environment Setup

We strongly recommend using a virtual environment:

### Using venv (Python 3.3+)

```bash
# Create virtual environment
python -m venv offitrans-env

# Activate (Linux/macOS)
source offitrans-env/bin/activate

# Activate (Windows)
offitrans-env\Scripts\activate

# Install Offitrans
pip install offitrans

# Install optional dependencies as needed
pip install openpyxl python-docx PyPDF2 python-pptx Pillow
```

### Using conda

```bash
# Create conda environment
conda create -n offitrans python=3.9

# Activate environment
conda activate offitrans

# Install from conda-forge (if available)
conda install -c conda-forge offitrans

# Or install with pip
pip install offitrans
```

## Verification

After installation, verify that Offitrans is working correctly:

```python
# Test basic import
import offitrans
print(f"Offitrans version: {offitrans.__version__}")

# Test translator
from offitrans import GoogleTranslator
translator = GoogleTranslator()
print("✅ Translator created successfully")

# Test available processors
from offitrans.processors import AVAILABLE_PROCESSORS
print(f"Available processors: {list(AVAILABLE_PROCESSORS.keys())}")

# Test specific processors based on installed dependencies
try:
    from offitrans import ExcelProcessor
    print("✅ Excel processing available")
except ImportError as e:
    print(f"❌ Excel processing not available: {e}")

try:
    from offitrans.processors import WordProcessor
    print("✅ Word processing available")
except ImportError as e:
    print(f"❌ Word processing not available: {e}")

try:
    from offitrans.processors import PDFProcessor
    print("✅ PDF processing available")
except ImportError as e:
    print(f"❌ PDF processing not available: {e}")

try:
    from offitrans.processors import PowerPointProcessor
    print("✅ PowerPoint processing available")
except ImportError as e:
    print(f"❌ PowerPoint processing not available: {e}")
```

## API Key Configuration

### Google Translate API

Offitrans works with both free and paid Google Translate APIs:

#### Free API (Default)
No configuration needed - works out of the box with rate limits.

#### Paid API (Google Cloud Translation)
1. Create a Google Cloud project
2. Enable the Translation API
3. Create an API key
4. Configure Offitrans:

```python
import os
from offitrans import GoogleTranslator

# SECURE: Use environment variable for API key
translator = GoogleTranslator(
    api_key=os.getenv('GOOGLE_TRANSLATE_API_KEY'),
    use_free_api=False
)
```

Or set environment variable:
```bash
# Set your actual API key here
export OFFITRANS_API_KEY="your-actual-google-api-key"
```

## Troubleshooting

### Common Issues

#### ImportError: No module named 'openpyxl'
```bash
pip install openpyxl
```

#### SSL Certificate Error
```bash
pip install --trusted-host pypi.org --trusted-host pypi.python.org offitrans
```

#### Permission Denied (Linux/macOS)
```bash
pip install --user offitrans
```

#### Out of Memory Errors
- Reduce `max_workers` in configuration
- Process smaller files or fewer files at once
- Increase system memory or use swap space

#### Translation API Errors
- Check your internet connection
- Verify API key (if using paid API)
- Check for rate limiting
- Try reducing `max_workers` to avoid hitting rate limits

### Getting Help

If you encounter issues:

1. **Check the logs**: Enable detailed logging to see what's happening
   ```python
   import logging
   logging.basicConfig(level=logging.DEBUG)
   ```

2. **Search existing issues**: Check [GitHub Issues](https://github.com/minglu6/Offitrans/issues)

3. **Create a new issue**: If you can't find a solution, create a detailed issue report

4. **Community support**: Join discussions on GitHub

## Performance Optimization

### For Large Files
```python
from offitrans.core.config import Config

config = Config()
config.translator.max_workers = 2  # Reduce for large files
config.cache.enabled = True  # Enable caching
config.processor.image_protection = False  # Disable if not needed
```

### For Many Small Files
```python
config = Config()
config.translator.max_workers = 10  # Increase for many small files
config.cache.enabled = True
config.cache.auto_save_interval = 20  # Save cache less frequently
```

### Memory Usage
- Monitor memory usage with large Excel files
- Consider processing files in batches
- Use streaming where possible

## Next Steps

After successful installation:

1. Read the [Quick Start Guide](quickstart.md)
2. Try the [Basic Usage Examples](../examples/basic_usage.py)
3. Explore the [API Documentation](api.md)
4. Check out [Advanced Examples](../examples/)