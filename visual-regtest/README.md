# Visual Regression Test

Automated visual regression testing for the `opc` and `docx` roundtrip layers.
Verifies that opening a `.docx` and saving it back produces a visually identical
document.

Two layers can be tested:

| `LAYER` | Roundtrip path                            | What it exercises                              |
|---------|-------------------------------------------|------------------------------------------------|
| `opc`   | `opc.OpenFile` → `pkg.SaveToFile`         | ZIP / OPC packaging only                       |
| `docx`  | `docx.OpenFile` → `doc.SaveFile`          | Full stack: OPC + XML parsing + typed parts    |

## Quick start

```bash
# 1. Drop your .docx files into the test-files/ directory
cp /path/to/your/*.docx visual-regtest/test-files/

# 2. Run (from the visual-regtest/ directory)
cd visual-regtest

make run                # OPC layer (default)
make run LAYER=docx     # DOCX layer
make run LAYER=opc      # OPC layer (explicit)

# 3. Open the report
make report
# or manually: open report/index.html
```

That's it. Drop files, pick a layer, `make run`.

## How it works

```
┌──────────────┐   Go roundtrip    ┌───────────────┐
│ original.docx │──(opc or docx)─▶ │ roundtrip.docx │
└──────┬────────┘                  └──────┬─────────┘
       │  LibreOffice + pdftoppm          │
       ▼                                  ▼
   page PNGs                          page PNGs
       │                                  │
       └──────────┐    ┌─────────────────┘
                  ▼    ▼
              SSIM comparison
                    │
                    ▼
             HTML report with
          thumbnails & diff maps
```

1. **Go roundtrip** — OPC or DOCX layer open → save (parallel workers)
2. **LibreOffice headless** — original & roundtripped `.docx` → PDF
3. **pdftoppm** — each PDF page → PNG
4. **SSIM** — per-page structural similarity (scikit-image)
5. **Report** — `report/index.html` with side-by-side thumbnails, diff heatmaps, scores

Everything runs inside a single Docker container — no local dependencies needed.

## Requirements

- Docker (with BuildKit)

## Configuration

All optional. Pass as make variables:

```bash
make run LAYER=docx SSIM_THRESHOLD=0.95 DPI=200 WORKERS=8
```

| Variable         | Default | Description                              |
|------------------|---------|------------------------------------------|
| `LAYER`          | `opc`   | Layer to test: `opc` or `docx`           |
| `SSIM_THRESHOLD` | `0.98`  | Minimum acceptable SSIM score            |
| `DPI`            | `150`   | Rendering resolution for page images     |
| `WORKERS`        | `4`     | Parallel worker count                    |

## Make targets

| Target   | Description                                    |
|----------|------------------------------------------------|
| `run`    | Build image + run full pipeline (LAYER=opc\|docx) |
| `build`  | Build Docker image only                        |
| `report` | Open the HTML report in a browser              |
| `clean`  | Remove report dir and Docker image             |
| `help`   | Show available targets                         |

## Report output

```
report/
├── index.html          # main report — open in browser
├── index.json          # machine-readable results for CI
└── images/
    └── <docx-stem>/
        ├── orig-1.png  # original page rendering
        ├── rt-1.png    # roundtripped page rendering
        └── diff-1.png  # SSIM difference heatmap
```

## CI integration

The container exits with code 0 always (report is the artifact).
Parse `report/index.json` programmatically, or check stderr for the summary line.

## File layout

```
visual-regtest/
├── test-files/             ← put your .docx files here
│   └── .gitkeep
├── report/                 ← generated (gitignored)
├── roundtrip/
│   ├── opc/main.go         # OPC-only:  opc.OpenFile → SaveToFile
│   └── docx/main.go        # Full docx: docx.OpenFile → doc.SaveFile
├── scripts/
│   ├── entrypoint.sh       # pipeline orchestrator (LAYER-aware)
│   └── compare_ssim.py     # SSIM comparison + HTML report
├── Dockerfile
├── docker-compose.yml
└── Makefile
```
