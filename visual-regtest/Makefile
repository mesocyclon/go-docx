# visual-regtest/Makefile
# ============================================================================
# Visual regression testing for OPC and DOCX roundtrip layers.
#
# 1. Put your .docx files into visual-regtest/test-files/
# 2. Run:  make run              (DOCX layer, default)
#     or:  make run LAYER=opc   (opc layer)
# 3. Open: visual-regtest/report/index.html
#
# Optional overrides:
#   make run LAYER=docx SSIM_THRESHOLD=0.95 DPI=200 WORKERS=8
# ============================================================================

SHELL      := /bin/bash
MAKEFLAGS  += --no-print-directory

# Defaults
LAYER          ?= docx
SSIM_THRESHOLD ?= 0.98
DPI            ?= 150
WORKERS        ?= 4
IMAGE_NAME     ?= visual-regtest
REPORT_DIR     := $(CURDIR)/report

# Project root = one level up from this Makefile.
PROJECT_ROOT := $(abspath $(CURDIR)/..)

.PHONY: build run clean report help check-files gen-files gen-clean replace-txt replace-txt-clean replace-tbl replace-tbl-clean replace-content replace-content-clean replace-user-mark replace-user-mark-clean replace-mark-batch-build replace-mark-batch-run replace-mark-batch-clean

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

check-files:
	@if [ -z "$$(ls -A $(CURDIR)/test-files/*.docx 2>/dev/null)" ]; then \
		echo ""; \
		echo "  ERROR: no .docx files in visual-regtest/test-files/"; \
		echo "  Put your test documents there and re-run."; \
		echo ""; \
		exit 1; \
	fi
	@echo "[make] $$(ls $(CURDIR)/test-files/*.docx | wc -l) .docx files in test-files/"

build: check-files ## Build the Docker image (includes test files)
	docker build \
		-t $(IMAGE_NAME) \
		-f $(CURDIR)/Dockerfile \
		$(PROJECT_ROOT)

run: build ## Run the full pipeline (LAYER=opc|docx)
	@mkdir -p $(REPORT_DIR)
	docker run --rm \
		-v "$(REPORT_DIR):/output" \
		--tmpfs /data:size=4G \
		-e LAYER=$(LAYER) \
		-e SSIM_THRESHOLD=$(SSIM_THRESHOLD) \
		-e DPI=$(DPI) \
		-e WORKERS=$(WORKERS) \
		$(IMAGE_NAME)
	@echo ""
	@echo "  Report: $(REPORT_DIR)/index.html"

clean: ## Remove report and Docker image
	rm -rf $(REPORT_DIR)
	-docker rmi $(IMAGE_NAME) 2>/dev/null

GEN_OUTPUT    := $(CURDIR)/gen-files/out

gen-files: ## Generate .docx test files from scratch using the public API
	@mkdir -p $(GEN_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/gen-files --output $(GEN_OUTPUT)
	@echo ""
	@echo "  Generated files: $(GEN_OUTPUT)/"
	@echo "  Manifest:        $(GEN_OUTPUT)/manifest.json"

gen-clean: ## Remove generated test files
	rm -rf $(GEN_OUTPUT)

REPLACE_TXT_OUTPUT := $(CURDIR)/replace-txt/out

replace-txt: ## Generate before/after .docx files for ReplaceText visual check
	@mkdir -p $(REPLACE_TXT_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/replace-txt --output $(REPLACE_TXT_OUTPUT)
	@echo ""
	@echo "  Before: $(REPLACE_TXT_OUTPUT)/01_before_replace.docx"
	@echo "  After:  $(REPLACE_TXT_OUTPUT)/02_after_replace.docx"

replace-txt-clean: ## Remove replace-txt output
	rm -rf $(REPLACE_TXT_OUTPUT)

REPLACE_TBL_OUTPUT := $(CURDIR)/replace-tbl/out

replace-tbl: ## Generate before/after .docx files for ReplaceWithTable visual check
	@mkdir -p $(REPLACE_TBL_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/replace-tbl --output $(REPLACE_TBL_OUTPUT)
	@echo ""
	@echo "  Before: $(REPLACE_TBL_OUTPUT)/01_before_replace_tbl.docx"
	@echo "  After:  $(REPLACE_TBL_OUTPUT)/02_after_replace_tbl.docx"

replace-tbl-clean: ## Remove replace-tbl output
	rm -rf $(REPLACE_TBL_OUTPUT)

REPLACE_CONTENT_OUTPUT := $(CURDIR)/replace-content/out

replace-content: ## Generate template, sources and result .docx files for ReplaceWithContent visual check
	@mkdir -p $(REPLACE_CONTENT_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/replace-content --output $(REPLACE_CONTENT_OUTPUT)
	@echo ""
	@echo "  Template: $(REPLACE_CONTENT_OUTPUT)/01_template.docx"
	@echo "  Sources:  $(REPLACE_CONTENT_OUTPUT)/sources/"
	@echo "  Result:   $(REPLACE_CONTENT_OUTPUT)/02_result.docx"

replace-content-clean: ## Remove replace-content output
	rm -rf $(REPLACE_CONTENT_OUTPUT)

REPLACE_USER_MARK_INPUT  := $(CURDIR)/replace-user-mark/in
REPLACE_USER_MARK_OUTPUT := $(CURDIR)/replace-user-mark/out

replace-user-mark: ## Fill user template (in/) with mark documents, save to out/
	@mkdir -p $(REPLACE_USER_MARK_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/replace-user-mark \
		--input $(REPLACE_USER_MARK_INPUT) \
		--output $(REPLACE_USER_MARK_OUTPUT)
	@echo ""
	@echo "  Input:    $(REPLACE_USER_MARK_INPUT)/template.docx"
	@echo "  Marks:    $(REPLACE_USER_MARK_INPUT)/mark/*.docx"
	@echo "  Result:   $(REPLACE_USER_MARK_OUTPUT)/result.docx"

replace-user-mark-clean: ## Remove replace-user-mark output
	rm -rf $(REPLACE_USER_MARK_OUTPUT)

# ── replace-user-mark-batch ──────────────────────────────────────────────
# For each .docx in replace-user-mark-batch/in/mark/, fills the template
# and roundtrips the result through the full Docker SSIM pipeline.
#
#   make replace-mark-batch-run                          # all variants
#   make replace-mark-batch-run VARIANT=keep_src         # single variant
#   make replace-mark-batch-run TAG='(<mytag>)' SSIM_THRESHOLD=0.95
#
# Available variants:
#   use_dest, keep_src, keep_src_force,
#   keep_diff, keep_diff_force, keep_num
# ─────────────────────────────────────────────────────────────────────────

TAG                        ?= (<mark>)
VARIANT                    ?=
REPLACE_MARK_BATCH_IMAGE   ?= replace-mark-batch-regtest
REPLACE_MARK_BATCH_REPORT  := $(CURDIR)/replace-user-mark-batch/report
REPLACE_MARK_BATCH_OUT     := $(CURDIR)/replace-user-mark-batch/out

replace-mark-batch-check:
	@if [ ! -f "$(CURDIR)/replace-user-mark-batch/in/template.docx" ]; then \
		echo ""; \
		echo "  ERROR: template not found in visual-regtest/replace-user-mark-batch/in/"; \
		echo "  Place template.docx with $(TAG) placeholders there."; \
		echo ""; \
		exit 1; \
	fi
	@NMARKS=$$(find "$(CURDIR)/replace-user-mark-batch/in/mark" -maxdepth 1 -iname '*.docx' 2>/dev/null | wc -l); \
	if [ "$$NMARKS" -eq 0 ]; then \
		echo ""; \
		echo "  ERROR: no .docx files in visual-regtest/replace-user-mark-batch/in/mark/"; \
		echo "  Place mark documents there."; \
		echo ""; \
		exit 1; \
	fi; \
	echo "[make] $$NMARKS mark file(s) in replace-user-mark-batch/in/mark/"

replace-mark-batch-build: replace-mark-batch-check ## Build Docker image for replace-mark-batch
	docker build \
		-t $(REPLACE_MARK_BATCH_IMAGE) \
		-f $(CURDIR)/Dockerfile \
		$(PROJECT_ROOT)

replace-mark-batch-run: replace-mark-batch-build ## Run replace-mark-batch SSIM pipeline (VARIANT= for single)
	@mkdir -p $(REPLACE_MARK_BATCH_REPORT) $(REPLACE_MARK_BATCH_OUT)
	docker run --rm \
		-v "$(REPLACE_MARK_BATCH_REPORT):/output" \
		-v "$(REPLACE_MARK_BATCH_OUT):/results" \
		--tmpfs /data:size=4G \
		-e LAYER=replace-mark-batch \
		-e SSIM_THRESHOLD=$(SSIM_THRESHOLD) \
		-e DPI=$(DPI) \
		-e WORKERS=$(WORKERS) \
		-e 'TAG=$(TAG)' \
		-e 'VARIANT=$(VARIANT)' \
		$(REPLACE_MARK_BATCH_IMAGE)
	@echo ""
	@echo "  Results: $(REPLACE_MARK_BATCH_OUT)/"
	@echo "  Reports: $(REPLACE_MARK_BATCH_REPORT)/<variant>/index.html"

replace-mark-batch-clean: ## Remove replace-mark-batch output, report, and image
	rm -rf $(REPLACE_MARK_BATCH_OUT) $(REPLACE_MARK_BATCH_REPORT)
	-docker rmi $(REPLACE_MARK_BATCH_IMAGE) 2>/dev/null

report: ## Open the HTML report (macOS / Linux)
	@if [ -f "$(REPORT_DIR)/index.html" ]; then \
		command -v xdg-open >/dev/null && xdg-open "$(REPORT_DIR)/index.html" || \
		command -v open >/dev/null && open "$(REPORT_DIR)/index.html" || \
		echo "Open $(REPORT_DIR)/index.html in your browser"; \
	else \
		echo "No report found. Run 'make run' first."; \
	fi