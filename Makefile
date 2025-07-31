# Makefile for Offitrans development

.PHONY: help install install-dev test test-cov lint format clean build upload docs

help:
	@echo "Available commands:"
	@echo "  install      Install package in development mode"
	@echo "  install-dev  Install package and development dependencies"
	@echo "  test         Run tests"
	@echo "  test-cov     Run tests with coverage"
	@echo "  lint         Run linting checks"
	@echo "  format       Format code with black and isort"
	@echo "  clean        Clean build artifacts"
	@echo "  build        Build package"
	@echo "  upload       Upload package to PyPI"
	@echo "  docs         Build documentation"

install:
	pip install -e .

install-dev:
	pip install -e ".[dev,full]"
	pre-commit install

test:
	pytest tests/ -v

test-cov:
	pytest tests/ -v --cov=offitrans --cov-report=term-missing --cov-report=html

test-unit:
	pytest tests/ -v -m "not integration and not slow"

test-integration:
	pytest tests/ -v -m "integration"

test-all:
	pytest tests/ -v --cov=offitrans --cov-report=term-missing

lint:
	flake8 offitrans tests examples
	mypy offitrans
	bandit -r offitrans

format:
	black offitrans tests examples
	isort offitrans tests examples

format-check:
	black --check offitrans tests examples
	isort --check-only offitrans tests examples

pre-commit:
	pre-commit run --all-files

clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf .pytest_cache/
	rm -rf .coverage
	rm -rf htmlcov/
	rm -rf .mypy_cache/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

build: clean
	python -m build

upload: build
	python -m twine upload dist/*

upload-test: build
	python -m twine upload --repository testpypi dist/*

docs:
	cd docs && make html

# Development workflow shortcuts
dev-setup: install-dev
	@echo "Development environment set up successfully!"

dev-test: format-check lint test-unit
	@echo "Development tests passed!"

ci-test: format-check lint test-all
	@echo "CI tests passed!"

# Quick development cycle
quick: format test-unit
	@echo "Quick development cycle completed!"

# Release preparation
release-check: clean format-check lint test-all build
	@echo "Release check completed successfully!"