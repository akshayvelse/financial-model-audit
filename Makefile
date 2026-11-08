.PHONY: all install test run clean

PYTHON = python
PYTEST = pytest

all: install run test

install:
	$(PYTHON) -m pip install -r requirements.txt

run:
	$(PYTHON) src/generate_models.py
	$(PYTHON) src/formula_scanner.py
	$(PYTHON) src/valuation_bridge.py

test:
	$(PYTEST) -v

clean:
	rm -rf __pycache__ .pytest_cache
