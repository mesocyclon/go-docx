#!/usr/bin/env python3
"""
compare_ssim.py – Compare original vs roundtripped .docx renderings using SSIM.

Workflow:
  1. Convert .docx → PDF via LibreOffice (batch).
  2. Convert PDF → page PNGs via pdftoppm (poppler).
  3. Compute per-page SSIM between original and roundtripped PNGs.
  4. Produce an HTML report with thumbnails and scores.

Usage:
  python3 compare_ssim.py \
      --original-dir /data/original \
      --roundtrip-dir /data/roundtrip \
      --work-dir /data/work \
      --report /data/report/index.html \
      --threshold 0.98 \
      --workers 4
"""

import argparse
import gc
import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

# Allow very large images — some docx pages render to 100M+ pixels.
# Without this, PIL raises DecompressionBombError and kills the worker.
Image.MAX_IMAGE_PIXELS = None

# Cap for SSIM comparison: downscale images larger than this to save RAM.
# 4000px per side is plenty for SSIM accuracy at 150 DPI.
_MAX_SSIM_SIDE = 4000


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------

def docx_to_pdf_batch(docx_dir: Path, pdf_dir: Path) -> None:
    """Convert all .docx in *docx_dir* to PDF using LibreOffice headless.

    Resilient: if a batch chunk fails (e.g. encrypted/corrupt file), falls
    back to converting each file in that chunk individually so one bad file
    doesn't kill the entire run.
    """
    pdf_dir.mkdir(parents=True, exist_ok=True)
    base_cmd = [
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(pdf_dir),
    ]
    docx_files = sorted(docx_dir.glob("*.docx"))
    if not docx_files:
        return

    chunk_size = 50
    for i in range(0, len(docx_files), chunk_size):
        chunk = docx_files[i:i + chunk_size]
        result = subprocess.run(
            base_cmd + [str(f) for f in chunk],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            timeout=600,
        )
        if result.returncode != 0:
            # Batch failed — retry each file individually so only the
            # broken ones are skipped.
            for f in chunk:
                r = subprocess.run(
                    base_cmd + [str(f)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.PIPE,
                    timeout=120,
                )
                if r.returncode != 0:
                    print(f"  [warn] LibreOffice failed on {f.name}, skipping")


def pdf_to_pngs(pdf_path: Path, png_dir: Path, dpi: int = 150) -> list[Path]:
    """Render each page of *pdf_path* as a PNG using pdftoppm."""
    png_dir.mkdir(parents=True, exist_ok=True)
    stem = pdf_path.stem
    prefix = png_dir / stem

    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        timeout=120,
    )
    # pdftoppm names output as <prefix>-1.png, <prefix>-2.png, …
    pages = sorted(png_dir.glob(f"{stem}-*.png"))
    # Fallback for single-page docs that produce <prefix>-1.png
    if not pages:
        pages = sorted(png_dir.glob(f"{stem}*.png"))
    return pages


# ---------------------------------------------------------------------------
# SSIM comparison
# ---------------------------------------------------------------------------

def _load_gray(path: Path) -> np.ndarray:
    """Load image as grayscale numpy array, downscaling if too large."""
    img = Image.open(path).convert("L")
    w, h = img.size
    if w > _MAX_SSIM_SIDE or h > _MAX_SSIM_SIDE:
        scale = _MAX_SSIM_SIDE / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    return np.array(img)


def compute_ssim(img_a_path: Path, img_b_path: Path) -> tuple[float, np.ndarray | None]:
    """Return (ssim_score, diff_image_array) for two page images."""
    a = _load_gray(img_a_path)
    b = _load_gray(img_b_path)

    # Resize to match if shapes differ (shouldn't happen, but safety).
    if a.shape != b.shape:
        h = max(a.shape[0], b.shape[0])
        w = max(a.shape[1], b.shape[1])
        a_pad = np.ones((h, w), dtype=np.uint8) * 255
        b_pad = np.ones((h, w), dtype=np.uint8) * 255
        a_pad[: a.shape[0], : a.shape[1]] = a
        b_pad[: b.shape[0], : b.shape[1]] = b
        a, b = a_pad, b_pad

    score, diff = ssim(a, b, full=True)
    # diff is float64 [0..1]; convert to uint8 for saving
    diff_img = (255 * (1 - diff)).clip(0, 255).astype(np.uint8)
    return float(score), diff_img


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class PageResult:
    page: int
    ssim_score: float
    orig_png: str
    rt_png: str
    diff_png: str


@dataclass
class FileReport:
    name: str
    ok: bool = True
    error: str = ""
    pages: list[PageResult] = field(default_factory=list)
    min_ssim: float = 1.0
    mean_ssim: float = 1.0


# ---------------------------------------------------------------------------
# Per-file worker
# ---------------------------------------------------------------------------

def process_one_file(
    name: str,
    orig_pdf_dir: Path,
    rt_pdf_dir: Path,
    work_dir: Path,
    report_img_dir: Path,
    dpi: int,
) -> FileReport:
    stem = Path(name).stem
    report = FileReport(name=name)

    try:
        return _process_one_file_inner(name, stem, report, orig_pdf_dir, rt_pdf_dir, work_dir, report_img_dir, dpi)
    except MemoryError:
        report.ok = False
        report.error = "out of memory (image too large)"
        return report
    except Exception as exc:
        report.ok = False
        report.error = str(exc)
        return report


def _process_one_file_inner(
    name: str,
    stem: str,
    report: FileReport,
    orig_pdf_dir: Path,
    rt_pdf_dir: Path,
    work_dir: Path,
    report_img_dir: Path,
    dpi: int,
) -> FileReport:

    orig_pdf = orig_pdf_dir / f"{stem}.pdf"
    rt_pdf = rt_pdf_dir / f"{stem}.pdf"

    if not orig_pdf.exists():
        report.ok = False
        report.error = "original PDF missing (LibreOffice conversion likely failed)"
        return report
    if not rt_pdf.exists():
        report.ok = False
        report.error = "roundtrip PDF missing (LibreOffice conversion likely failed)"
        return report

    orig_png_dir = work_dir / "orig_png" / stem
    rt_png_dir = work_dir / "rt_png" / stem
    diff_png_dir = report_img_dir / stem

    try:
        orig_pages = pdf_to_pngs(orig_pdf, orig_png_dir, dpi)
        rt_pages = pdf_to_pngs(rt_pdf, rt_png_dir, dpi)
    except subprocess.CalledProcessError as exc:
        report.ok = False
        report.error = f"pdftoppm failed: {exc}"
        return report

    if not orig_pages:
        report.ok = False
        report.error = "original PDF produced no pages"
        return report
    if not rt_pages:
        report.ok = False
        report.error = "roundtrip PDF produced no pages"
        return report

    # Page count mismatch → report but still compare overlapping pages.
    max_pages = max(len(orig_pages), len(rt_pages))

    diff_png_dir.mkdir(parents=True, exist_ok=True)
    scores = []
    for idx in range(max_pages):
        if idx >= len(orig_pages) or idx >= len(rt_pages):
            report.pages.append(PageResult(
                page=idx + 1,
                ssim_score=0.0,
                orig_png=str(orig_pages[idx]) if idx < len(orig_pages) else "",
                rt_png=str(rt_pages[idx]) if idx < len(rt_pages) else "",
                diff_png="",
            ))
            scores.append(0.0)
            continue

        score, diff_arr = compute_ssim(orig_pages[idx], rt_pages[idx])
        scores.append(score)

        # Save diff image.
        diff_path = diff_png_dir / f"diff-{idx + 1}.png"
        Image.fromarray(diff_arr).save(diff_path)

        # Copy originals + rt into report img dir for the HTML report.
        orig_dst = diff_png_dir / f"orig-{idx + 1}.png"
        rt_dst = diff_png_dir / f"rt-{idx + 1}.png"
        shutil.copy2(orig_pages[idx], orig_dst)
        shutil.copy2(rt_pages[idx], rt_dst)

        report.pages.append(PageResult(
            page=idx + 1,
            ssim_score=score,
            orig_png=str(orig_dst),
            rt_png=str(rt_dst),
            diff_png=str(diff_path),
        ))

    report.min_ssim = min(scores) if scores else 1.0
    report.mean_ssim = float(np.mean(scores)) if scores else 1.0

    # Clean up temporary full-size PNGs to free tmpfs space.
    shutil.rmtree(orig_png_dir, ignore_errors=True)
    shutil.rmtree(rt_png_dir, ignore_errors=True)

    return report


# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------

def _img_tag(path_str: str, report_dir: Path, width: int = 320) -> str:
    """Return an <img> tag with a relative path from the report dir."""
    if not path_str:
        return "<em>missing</em>"
    try:
        rel = os.path.relpath(path_str, report_dir)
    except ValueError:
        rel = path_str
    return f'<img src="{rel}" width="{width}" loading="lazy">'


def generate_html_report(
    reports: list[FileReport],
    report_path: Path,
    threshold: float,
) -> None:
    report_dir = report_path.parent

    total = len(reports)
    ok_count = sum(1 for r in reports if r.ok)
    fail_count = total - ok_count
    below = sum(1 for r in reports if r.ok and r.min_ssim < threshold)
    perfect = sum(1 for r in reports if r.ok and r.min_ssim >= threshold)

    # Sort: failures first, then by min_ssim ascending.
    reports_sorted = sorted(reports, key=lambda r: (r.ok, r.min_ssim))

    rows = []
    for r in reports_sorted:
        status_cls = "fail" if not r.ok else ("warn" if r.min_ssim < threshold else "pass")
        status_text = r.error if not r.ok else f"min={r.min_ssim:.4f}  mean={r.mean_ssim:.4f}"

        detail = ""
        if r.ok and r.pages:
            page_rows = []
            for p in r.pages:
                pcls = "warn" if p.ssim_score < threshold else "pass"
                page_rows.append(f"""
                <tr class="{pcls}">
                  <td>Page {p.page}</td>
                  <td>{p.ssim_score:.4f}</td>
                  <td>{_img_tag(p.orig_png, report_dir)}</td>
                  <td>{_img_tag(p.rt_png, report_dir)}</td>
                  <td>{_img_tag(p.diff_png, report_dir)}</td>
                </tr>""")

            detail = f"""
            <details>
              <summary>Pages ({len(r.pages)})</summary>
              <table class="pages">
                <tr><th>Page</th><th>SSIM</th><th>Original</th><th>Roundtrip</th><th>Diff</th></tr>
                {"".join(page_rows)}
              </table>
            </details>"""

        rows.append(f"""
        <tr class="{status_cls}">
          <td class="fname">{r.name}</td>
          <td>{status_text}</td>
        </tr>
        <tr><td colspan="2">{detail}</td></tr>
        """)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>OPC Visual Regression Report</title>
<style>
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 2rem; background: #fafafa; }}
  h1 {{ color: #1a1a2e; }}
  .summary {{ display: flex; gap: 2rem; margin-bottom: 2rem; }}
  .summary .card {{ background: #fff; border-radius: 8px; padding: 1rem 2rem; box-shadow: 0 1px 3px rgba(0,0,0,.1); }}
  .card .num {{ font-size: 2rem; font-weight: 700; }}
  .card.ok .num {{ color: #16a34a; }}
  .card.warn .num {{ color: #d97706; }}
  .card.err .num {{ color: #dc2626; }}
  table {{ border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.1); }}
  th, td {{ padding: .6rem 1rem; text-align: left; border-bottom: 1px solid #eee; }}
  th {{ background: #1a1a2e; color: #fff; }}
  .pass {{ background: #f0fdf4; }}
  .warn {{ background: #fffbeb; }}
  .fail {{ background: #fef2f2; }}
  .fname {{ font-family: monospace; }}
  details {{ margin: .3rem 0; }}
  table.pages img {{ border: 1px solid #ddd; border-radius: 4px; }}
  table.pages td {{ vertical-align: top; }}
</style>
</head>
<body>
<h1>OPC Visual Regression Report</h1>
<div class="summary">
  <div class="card ok"><div class="num">{perfect}</div>Pass (SSIM &ge; {threshold})</div>
  <div class="card warn"><div class="num">{below}</div>Below threshold</div>
  <div class="card err"><div class="num">{fail_count}</div>Errors</div>
  <div class="card"><div class="num">{total}</div>Total files</div>
</div>
<table>
  <tr><th>File</th><th>Result</th></tr>
  {"".join(rows)}
</table>
</body>
</html>"""

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(html, encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="SSIM visual regression comparison")
    parser.add_argument("--original-dir", required=True, help="dir with original .docx")
    parser.add_argument("--roundtrip-dir", required=True, help="dir with roundtripped .docx")
    parser.add_argument("--work-dir", required=True, help="scratch space for PDFs/PNGs")
    parser.add_argument("--report", required=True, help="output HTML report path")
    parser.add_argument("--threshold", type=float, default=0.98, help="SSIM pass threshold")
    parser.add_argument("--dpi", type=int, default=150, help="rendering DPI")
    parser.add_argument("--workers", type=int, default=4, help="parallel workers")
    args = parser.parse_args()

    original_dir = Path(args.original_dir)
    roundtrip_dir = Path(args.roundtrip_dir)
    work_dir = Path(args.work_dir)
    report_path = Path(args.report)
    report_img_dir = report_path.parent / "images"

    # ---- read manifest from roundtrip tool ----
    manifest_path = roundtrip_dir / "manifest.json"
    if manifest_path.exists():
        with open(manifest_path) as f:
            manifest = json.load(f)
        file_names = [e["name"] for e in manifest]
        failed_rt = {e["name"] for e in manifest if not e["ok"]}
    else:
        file_names = [p.name for p in sorted(original_dir.glob("*.docx"))]
        failed_rt = set()

    print(f"[compare] {len(file_names)} files to compare")

    # ---- step 1: convert .docx → .pdf (batch via LibreOffice) ----
    orig_pdf_dir = work_dir / "orig_pdf"
    rt_pdf_dir = work_dir / "rt_pdf"

    print("[compare] converting originals to PDF …")
    docx_to_pdf_batch(original_dir, orig_pdf_dir)

    print("[compare] converting roundtripped to PDF …")
    docx_to_pdf_batch(roundtrip_dir, rt_pdf_dir)

    # ---- step 2: per-file SSIM comparison ----
    reports: list[FileReport] = []

    # Pre-populate roundtrip failures.
    for name in file_names:
        if name in failed_rt:
            reports.append(FileReport(name=name, ok=False, error="roundtrip failed (Go)"))

    compare_names = [n for n in file_names if n not in failed_rt]
    print(f"[compare] comparing {len(compare_names)} files …")

    for i, name in enumerate(compare_names, 1):
        try:
            rep = process_one_file(
                name, orig_pdf_dir, rt_pdf_dir, work_dir, report_img_dir, args.dpi,
            )
        except Exception as exc:
            rep = FileReport(name=name, ok=False, error=str(exc))
        reports.append(rep)

        # Free memory after each file — large PNGs can eat GBs.
        gc.collect()

        if i % 50 == 0 or i == len(compare_names):
            print(f"  [{i}/{len(compare_names)}]")

    # ---- step 3: generate report ----
    generate_html_report(reports, report_path, args.threshold)

    # ---- also dump JSON for CI ----
    json_path = report_path.with_suffix(".json")
    json_path.write_text(json.dumps([asdict(r) for r in reports], indent=2), encoding="utf-8")

    below = [r for r in reports if r.ok and r.min_ssim < args.threshold]
    errors = [r for r in reports if not r.ok]
    print(f"\n[compare] DONE — {len(reports)} files, "
          f"{len(reports) - len(errors) - len(below)} pass, "
          f"{len(below)} below threshold, {len(errors)} errors")
    print(f"[compare] report: {report_path}")

    # Exit with failure code if anything is below threshold.
    if below or errors:
        sys.exit(1)


if __name__ == "__main__":
    main()