# Contributing Guide

Thank you for your interest in the Offitrans project! We welcome all forms of contributions, including but not limited to:

- 🐛 Bug reports
- 💡 Feature suggestions
- 📝 Documentation improvements
- 🔧 Code fixes
- ✨ New features
- 🌍 Translation and localization

[中文贡献指南](CONTRIBUTING_CN.md)

## 📋 Getting Started

### Development Environment Setup

1. **Fork the Project**
   ```bash
   # Fork the project on GitHub to your account
   # Then clone your fork
   git clone https://github.com/your-username/Offitrans.git
   cd Offitrans
   ```

2. **Set Up Development Environment**
   ```bash
   # Create virtual environment
   python -m venv venv
   
   # Activate virtual environment
   # Windows
   venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   
   # Install dependencies
   pip install -e .[dev]
   
   # Install pre-commit hooks
   pre-commit install
   ```

3. **Create Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-fix-name
   ```

## 🐛 Reporting Bugs

If you find a bug, please report it via [GitHub Issues](https://github.com/minglu6/Offitrans/issues).

**Bug reports should include:**

- 🔍 **Clear title and description**
- 📱 **Environment information** (Python version, OS, etc.)
- 📝 **Steps to reproduce**
- 🎯 **Expected vs actual behavior**
- 📋 **Relevant error logs or screenshots**
- 📄 **Sample files** (if specific Office files are involved)

### Bug Report Template

```markdown
## Bug Description
A clear and concise description of what the bug is.

## Steps to Reproduce
1. Go to '...'
2. Click on '....'
3. Scroll down to '....'
4. See error

## Expected Behavior
A clear and concise description of what you expected to happen.

## Actual Behavior
A clear and concise description of what actually happened.

## Environment
- OS: [e.g. Windows 10, macOS 12.1, Ubuntu 20.04]
- Python Version: [e.g. 3.9.7]
- Offitrans Version: [e.g. 1.0.0]

## Additional Context
Add any other context about the problem here.
```

## 💡 Feature Requests

We welcome feature suggestions! Please submit them via [GitHub Issues](https://github.com/minglu6/Offitrans/issues).

**Feature requests should include:**

- 🎯 **Problem description**: What problem are you trying to solve?
- 💡 **Proposed solution**: How would your suggested feature solve this problem?
- 🔄 **Alternative solutions**: Have you considered other solutions?
- 📊 **Use cases**: Who would use this feature and when?

## 🔧 Code Contributions

### Coding Standards

1. **Code Style**
   - Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/) coding standards
   - Use 4 spaces for indentation
   - Line length limit of 88 characters (Black default)

2. **Naming Conventions**
   - Class names use `PascalCase`
   - Function and variable names use `snake_case`
   - Constants use `UPPER_CASE`
   - Private methods and attributes start with `_`

3. **Docstrings**
   ```python
   def translate_text(self, text: str, target_language: str = 'en') -> str:
       """
       Translate text content.
       
       Args:
           text: Text to be translated
           target_language: Target language code
           
       Returns:
           Translated text
           
       Raises:
           ValueError: When input parameters are invalid
           TranslationError: When translation fails
       """
   ```

4. **Type Hints**
   - All public methods should have type hints
   - Use `typing` module for complex type definitions

### Code Quality Tools

Before submitting code, please check code quality with these tools:

```bash
# Code formatting
black .

# Import sorting
isort .

# Code style check
flake8 .

# Type checking
mypy .

# Run tests
pytest tests/ -v --cov=offitrans
```

### Commit Guidelines

Use [Conventional Commits](https://www.conventionalcommits.org/) specification:

- `feat:` New features
- `fix:` Bug fixes
- `docs:` Documentation updates
- `style:` Code formatting
- `refactor:` Code refactoring
- `test:` Test-related changes
- `chore:` Build process or auxiliary tool changes

**Examples:**
```bash
git commit -m "feat: add image protection for PDF translation"
git commit -m "fix: resolve Excel merged cell formatting issue"
git commit -m "docs: update API usage documentation"
```

### Pull Request Process

1. **Ensure Code Quality**
   - All tests pass
   - Code style checks pass
   - New features have corresponding tests

2. **Create Pull Request**
   - Provide clear title and description
   - Explain the reason and content of changes
   - Reference related issues if fixing bugs

3. **PR Description Template**
   ```markdown
   ## Type of Change
   - [ ] Bug fix
   - [ ] New feature
   - [ ] Documentation update
   - [ ] Code refactoring
   - [ ] Performance improvement
   
   ## Description
   Clear description of the changes in this PR
   
   ## Related Issue
   Fixes #123
   
   ## Testing
   - [ ] Added unit tests
   - [ ] All existing tests pass
   - [ ] Manual testing completed
   
   ## Checklist
   - [ ] Code follows project coding standards
   - [ ] Added necessary documentation and comments
   - [ ] All tests pass
   ```

## 🧪 Testing

### Running Tests

```bash
# Run all tests
pytest

# Run specific test file
pytest tests/unit/test_processors.py

# Run tests with coverage report
pytest --cov=offitrans --cov-report=html

# Run specific test types
pytest -m unit          # Unit tests only
pytest -m integration   # Integration tests only
```

### Writing Tests

- Each new feature should have corresponding tests
- Test files should be named `test_*.py`
- Test methods should be named `test_*`

```python
def test_translate_excel_basic():
    """Test basic Excel translation functionality."""
    processor = ExcelProcessor()
    translator = GoogleTranslator()
    
    # Test implementation here
    assert result is not None
```

## 📝 Documentation Contributions

### Documentation Types

- **API Documentation**: Docstrings in code
- **User Guide**: README.md and usage examples
- **Developer Documentation**: CONTRIBUTING.md and technical specifications

### Documentation Standards

- Use clear and concise language
- Provide practical code examples
- Keep documentation synchronized between languages

## 🌍 Internationalization Contributions

We welcome multilingual support contributions:

- Translate documentation to other languages
- Add support for new translation languages
- Improve existing language translation quality

## 🎯 Project Priorities

Current focus areas for the project:

1. **Stability Improvements** - Fix bugs in existing features
2. **Performance Optimization** - Improve translation speed and memory efficiency
3. **Format Preservation** - Enhance style preservation for various Office formats
4. **New Format Support** - Add support for more file formats
5. **Multiple Translation Engines** - Integrate more translation services

## 📞 Getting Help

If you have any questions or need help:

- 📝 Create a [GitHub Issue](https://github.com/minglu6/Offitrans/issues)
- 💬 Join [Discussions](https://github.com/minglu6/Offitrans/discussions)

## 🙏 Contributors

Thanks to all developers who have contributed to Offitrans!

<!-- Contributors list can be added here or use GitHub's contributors API -->

---

Thank you again for your contribution to the Offitrans project! 🚀