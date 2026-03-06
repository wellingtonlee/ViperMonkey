.PHONY: install-dev test lint lint-fix clean

install-dev:
	pip install -e ".[dev]"
	pre-commit install

test:
	python -m pytest tests/ -v

lint:
	ruff check vipermonkey/ tests/
	ruff format --check vipermonkey/ tests/

lint-fix:
	ruff check --fix vipermonkey/ tests/
	ruff format vipermonkey/ tests/

clean:
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	find . -type f -name "*.pyc" -delete 2>/dev/null || true
	rm -rf .pytest_cache .ruff_cache dist build
	find . -type d -name "*.egg-info" -exec rm -rf {} + 2>/dev/null || true
