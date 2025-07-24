# Contributing to Excel Analyzer

Thank you for your interest in contributing to Excel Analyzer! This document provides guidelines and information for contributors.

## 🚀 Getting Started

### Prerequisites
- Python 3.12 or higher
- Git
- Basic knowledge of Excel file structures

### Development Setup

1. **Fork and Clone**
   ```bash
   git clone https://github.com/your-username/excel-analyzer.git
   cd excel-analyzer
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   pip install -e .[dev]
   ```

4. **Install Pre-commit Hooks**
   ```bash
   pre-commit install
   ```

## 📋 Development Guidelines

### Code Style
- Follow PEP 8 style guidelines
- Use Black for code formatting
- Maximum line length: 88 characters
- Use type hints for all functions

### Code Quality
- Write docstrings in reStructuredText format
- Use logging instead of print statements
- Follow single responsibility principle
- Write testable, modular code

### Testing
- Write tests for new features
- Maintain test coverage above 80%
- Use pytest for testing
- Include both unit and integration tests

### Documentation
- Update README.md for user-facing changes
- Add docstrings for new functions
- Update CHANGELOG.md for all changes
- Keep examples up to date

## 🔧 Development Workflow

### 1. Create a Feature Branch
```bash
git checkout -b feature/your-feature-name
```

### 2. Make Changes
- Write your code
- Add tests
- Update documentation
- Format code with Black

### 3. Test Your Changes
```bash
# Run tests
pytest

# Check code quality
flake8
mypy excel_parser.py

# Format code
black .
```

### 4. Commit Your Changes
```bash
git add .
git commit -m "feat: add new feature description"
```

### 5. Push and Create Pull Request
```bash
git push origin feature/your-feature-name
```

## 📝 Commit Message Format

Use conventional commit format:
```
type(scope): description

[optional body]

[optional footer]
```

Types:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes
- `refactor`: Code refactoring
- `test`: Test changes
- `chore`: Maintenance tasks

## 🧪 Testing

### Running Tests
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=excel_parser --cov-report=html

# Run specific test file
pytest tests/test_parser.py

# Run specific test function
pytest tests/test_parser.py::test_analyze_workbook
```

### Test Structure
```
tests/
├── test_parser.py          # Core parser tests
├── test_extractor.py       # Extractor tests
├── test_dataframes.py      # DataFrame extraction tests
└── conftest.py            # Test configuration
```

## 📊 Code Quality Tools

### Pre-commit Hooks
The project uses pre-commit hooks to ensure code quality:
- Black (code formatting)
- Flake8 (linting)
- MyPy (type checking)
- Pre-commit (hook management)

### Manual Checks
```bash
# Format code
black .

# Lint code
flake8

# Type check
mypy excel_parser.py

# Run all quality checks
make quality
```

## 🎯 Areas for Contribution

### High Priority
- [ ] Add more comprehensive test coverage
- [ ] Improve pivot table analysis
- [ ] Add support for more Excel features
- [ ] Performance optimization for large files

### Medium Priority
- [ ] Add web interface
- [ ] Create more example scripts
- [ ] Improve error handling
- [ ] Add configuration options

### Low Priority
- [ ] Add support for other file formats
- [ ] Create plugins system
- [ ] Add batch processing GUI
- [ ] Performance benchmarking tools

## 🐛 Reporting Issues

### Bug Reports
When reporting bugs, please include:
- Python version
- Operating system
- Excel file type and version
- Steps to reproduce
- Expected vs actual behavior
- Error messages (if any)

### Feature Requests
For feature requests, please include:
- Use case description
- Expected functionality
- Examples of similar features
- Priority level

## 📞 Getting Help

- **Issues**: Use GitHub Issues for bugs and feature requests
- **Discussions**: Use GitHub Discussions for questions and ideas
- **Documentation**: Check the docs/ directory for detailed guides

## 📄 License

By contributing to Excel Analyzer, you agree that your contributions will be licensed under the MIT License.

## 🙏 Acknowledgments

Thank you to all contributors who help make Excel Analyzer better!

---

**Note**: This is a living document. Please suggest improvements or clarifications through issues or pull requests. 