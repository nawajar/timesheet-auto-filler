.PHONY: install run clean help

PYTHON = python3
VENV = venv
BIN = $(VENV)/bin

install: ## Create venv and install dependencies
	$(PYTHON) -m venv $(VENV)
	$(BIN)/pip install --upgrade pip
	$(BIN)/pip install -r requirements.txt

run: ## Run the timesheet filler
	$(BIN)/python fill_timesheet.py

clean: ## Remove generated output file
	rm -f timesheet_filled.xlsx

reset: ## Remove venv and output file
	rm -rf $(VENV)
	rm -f timesheet_filled.xlsx

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-10s\033[0m %s\n", $$1, $$2}'