# Example Makefile for Files Viewer testing

APP=app.py
PYTHON=python

.PHONY: run lint test clean help

help:
	@echo "Targets: run, lint, test, clean"

run:
	$(PYTHON) $(APP)

lint:
	$(PYTHON) -m py_compile $(APP)

test:
	$(PYTHON) -m pytest -q

clean:
	@echo "Nothing to clean in this sample"
