# Security Best Practices for Offitrans

This document outlines security best practices when using Offitrans, particularly regarding API key management and data protection.

## 🔐 API Key Security

### ⚠️ Critical: Never Hardcode API Keys

**❌ Never do this:**
```python
# INSECURE - Never hardcode API keys in source code
translator = GoogleTranslator(api_key="AIzaSyC8BgGgPiGgPa...")
```

**✅ Always do this:**
```python
# SECURE - Use environment variables
import os
translator = GoogleTranslator(api_key=os.getenv('GOOGLE_TRANSLATE_API_KEY'))
```

### Environment Variable Configuration

Set your API keys using environment variables:

```bash
# Option 1: Google Translate API key
export GOOGLE_TRANSLATE_API_KEY="your-actual-api-key"

# Option 2: Generic Offitrans API key
export OFFITRANS_API_KEY="your-actual-api-key"

# Option 3: Using .env file (with python-dotenv)
echo "GOOGLE_TRANSLATE_API_KEY=your-actual-api-key" > .env
```

### Secure Configuration Loading

```python
import os
from pathlib import Path
from offitrans import GoogleTranslator

# Method 1: Direct environment variable
api_key = os.getenv('GOOGLE_TRANSLATE_API_KEY')
if not api_key:
    raise ValueError("API key not found in environment variables")

translator = GoogleTranslator(api_key=api_key)

# Method 2: Using python-dotenv for development
try:
    from dotenv import load_dotenv
    load_dotenv()  # Load .env file
    api_key = os.getenv('GOOGLE_TRANSLATE_API_KEY')
except ImportError:
    api_key = os.getenv('GOOGLE_TRANSLATE_API_KEY')

translator = GoogleTranslator(api_key=api_key)
```

## 🛡️ Data Protection

### Input Validation

Always validate and sanitize input data:

```python
import re
from pathlib import Path

def validate_file_path(file_path: str) -> bool:
    """Validate file path for security"""
    path = Path(file_path)
    
    # Check if path is within allowed directories
    try:
        path.resolve().relative_to(Path.cwd())
    except ValueError:
        return False
    
    # Check file extension
    allowed_extensions = {'.xlsx', '.docx', '.pptx', '.pdf'}
    if path.suffix.lower() not in allowed_extensions:
        return False
    
    return True

# Use validation before processing
if validate_file_path(input_file):
    processor.process_file(input_file, output_file)
else:
    raise ValueError("Invalid or unsafe file path")
```

### Sensitive Data Handling

Be careful with sensitive content in documents:

```python
import logging

# Configure logging to avoid exposing sensitive data
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Never log actual content - only metadata
logger.info(f"Processing file: {Path(file_path).name}")  # Only filename
logger.info(f"Translated {len(texts)} text segments")    # Only counts
```

### Rate Limiting and Error Handling

```python
from offitrans.core.config import Config

# Configure conservative rate limiting
config = Config()
config.translator.max_workers = 2  # Limit concurrent requests
config.translator.timeout = 30     # Set reasonable timeout
config.translator.retry_count = 3  # Limit retry attempts

# Handle errors securely
try:
    result = translator.translate_text(text)
except Exception as e:
    # Log error without exposing sensitive details
    logger.error(f"Translation failed: {type(e).__name__}")
    # Don't log the actual text content or API response
```

## 🔒 Production Security

### CI/CD Pipeline Security

When using Offitrans in CI/CD pipelines:

```yaml
# GitHub Actions example
name: Translation
on: [push]
jobs:
  translate:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v4
    - name: Setup Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.9'
    - name: Install dependencies
      run: pip install offitrans
    - name: Run translation
      env:
        GOOGLE_TRANSLATE_API_KEY: ${{ secrets.GOOGLE_TRANSLATE_API_KEY }}
      run: python translate_docs.py
```

### Docker Security

```dockerfile
FROM python:3.9-slim

# Don't include API keys in the image
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

# Use environment variables at runtime
CMD ["python", "app.py"]
```

```bash
# Pass API key as environment variable when running container
docker run -e GOOGLE_TRANSLATE_API_KEY="$GOOGLE_TRANSLATE_API_KEY" my-app
```

### File System Security

```python
import tempfile
import os
from pathlib import Path

# Use temporary directories for processing
with tempfile.TemporaryDirectory() as temp_dir:
    temp_input = Path(temp_dir) / "input.xlsx"
    temp_output = Path(temp_dir) / "output.xlsx"
    
    # Copy input file to temp location
    import shutil
    shutil.copy2(input_file, temp_input)
    
    # Process file
    processor.process_file(str(temp_input), str(temp_output))
    
    # Copy result back
    shutil.copy2(temp_output, output_file)
    
    # Temp files are automatically cleaned up
```

## 🚨 Security Checklist

Before deploying Offitrans in production:

- [ ] **No API keys in source code** - All keys stored in environment variables
- [ ] **Input validation** - Validate all file paths and user inputs
- [ ] **Error handling** - Don't expose sensitive information in error messages
- [ ] **Logging security** - Don't log API keys, file contents, or sensitive data
- [ ] **Access controls** - Restrict file system access to necessary directories only
- [ ] **Rate limiting** - Configure appropriate rate limits for API calls
- [ ] **Dependency security** - Keep all dependencies updated
- [ ] **Network security** - Use HTTPS for all API communications
- [ ] **Temporary files** - Clean up temporary files after processing
- [ ] **Audit logging** - Log security-relevant events without sensitive details

## 🔍 Security Monitoring

Monitor your application for security issues:

```python
import logging
from datetime import datetime

# Security audit logging
security_logger = logging.getLogger('offitrans.security')
security_handler = logging.FileHandler('security.log')
security_logger.addHandler(security_handler)

def log_security_event(event_type: str, details: dict):
    """Log security-relevant events"""
    event = {
        'timestamp': datetime.utcnow().isoformat(),
        'event_type': event_type,
        'details': {k: v for k, v in details.items() if k not in ['api_key', 'content']}
    }
    security_logger.info(f"SECURITY: {event}")

# Usage
log_security_event('file_processed', {
    'filename': Path(file_path).name,
    'file_size': os.path.getsize(file_path),
    'success': True
})
```

## 📞 Reporting Security Issues

If you discover security vulnerabilities in Offitrans:

1. **Do not** open a public GitHub issue
2. Email security concerns to the maintainers privately
3. Provide detailed information about the vulnerability
4. Allow reasonable time for fixes before public disclosure

## 📚 Additional Resources

- [OWASP Application Security Verification Standard](https://owasp.org/www-project-application-security-verification-standard/)
- [Google Cloud API Key Best Practices](https://cloud.google.com/docs/authentication/api-keys#securing_an_api_key)
- [Environment Variable Security](https://12factor.net/config)
- [Python Security Guidelines](https://python-security.readthedocs.io/)

---

*Last updated: 2025-08-07*
*Version: 1.0*