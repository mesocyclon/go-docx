#!/usr/bin/env bash
# entrypoint.sh – orchestrates the visual regression pipeline inside Docker.
#
# The test corpus is baked into the image at /corpus.
# A single roundtrip binary is pre-built:
#   /usr/local/bin/roundtrip --layer=opc   (OPC layer)
#   /usr/local/bin/roundtrip --layer=docx  (docx layer)
#
# Report output goes to /output (bind-mounted from the host).
#
# Environment variables (all have sensible defaults):
#   LAYER           – which layer to test: opc | docx  (default: opc)
#   SSIM_THRESHOLD  – SSIM pass threshold  (default: 0.98)
#   DPI             – rendering resolution  (default: 150)
#   WORKERS         – parallel workers      (default: 4)
set -euo pipefail

LAYER="${LAYER:-opc}"
THRESHOLD="${SSIM_THRESHOLD:-0.98}"
DPI="${DPI:-150}"
WORKERS="${WORKERS:-4}"

# Validate LAYER.
case "${LAYER}" in
    opc|docx|replace-mark-batch) ;;
    *)
        echo "[entrypoint] ERROR: LAYER must be 'opc', 'docx', or 'replace-mark-batch', got '${LAYER}'"
        exit 1
        ;;
esac

TAG="${TAG:-(<mark>)}"
VARIANT="${VARIANT:-}"

ROUNDTRIP_BIN="/usr/local/bin/roundtrip"
LABEL=$(echo "${LAYER}" | tr '[:lower:]' '[:upper:]')

DATA="/data"
ORIG_DIR="/corpus"
RT_DIR="${DATA}/roundtrip"
WORK_DIR="${DATA}/work"
REPORT_DIR="/output"

# ==========================================================================
# replace-mark-batch: generate filled docs, compare them against mark files
#
# Original = mark file (expected content)
# Result   = filled template (actual output of ReplaceWithContent)
# ==========================================================================
if [ "${LAYER}" = "replace-mark-batch" ]; then
    MARK_BATCH_IN="/mark-batch-in"
    MARK_DIR="${MARK_BATCH_IN}/mark"
    GEN_DIR="${DATA}/generated"

    echo "=============================================="
    echo " REPLACE-MARK-BATCH Visual Regression Test"
    echo "=============================================="
    echo " Tag:       ${TAG}"
    echo " Variant:   ${VARIANT:-all}"
    echo " Threshold: ${THRESHOLD}"
    echo " DPI:       ${DPI}"
    echo " Workers:   ${WORKERS}"
    echo "=============================================="

    if [ ! -f "${MARK_BATCH_IN}/template.docx" ]; then
        echo "[entrypoint] ERROR: template not found at ${MARK_BATCH_IN}/template.docx"
        echo "[entrypoint] Put template.docx into visual-regtest/replace-user-mark-batch/in/"
        exit 1
    fi

    NMARKS=$(find "${MARK_DIR}" -maxdepth 1 -iname '*.docx' 2>/dev/null | wc -l)
    echo "[entrypoint] found ${NMARKS} mark files"

    if [ "${NMARKS}" -eq 0 ]; then
        echo "[entrypoint] ERROR: no .docx files in ${MARK_DIR}/"
        echo "[entrypoint] Put mark documents into visual-regtest/replace-user-mark-batch/in/mark/"
        exit 1
    fi

    # Step 1: Generate filled documents into per-variant subdirs.
    mkdir -p "${GEN_DIR}"
    echo "[entrypoint] generating filled documents …"
    VARIANT_FLAG=""
    if [ -n "${VARIANT}" ]; then
        VARIANT_FLAG="--variant=${VARIANT}"
    fi
    /usr/local/bin/replace-mark-batch \
        --input="${MARK_BATCH_IN}" \
        --output="${GEN_DIR}" \
        --tag="${TAG}" \
        --workers="${WORKERS}" \
        ${VARIANT_FLAG}

    # Copy results (preserving variant subdirs) to /results (bind-mounted).
    if [ -d /results ]; then
        cp -r "${GEN_DIR}"/*/ /results/ 2>/dev/null || true
        echo "[entrypoint] copied results to /results/"
    fi

    # Step 2: SSIM comparison per variant subdir.
    for VDIR in "${GEN_DIR}"/*/; do
        [ -d "${VDIR}" ] || continue
        VNAME=$(basename "${VDIR}")
        NGENERATED=$(find "${VDIR}" -maxdepth 1 -iname '*.docx' | wc -l)
        echo "[entrypoint] variant ${VNAME}: ${NGENERATED} .docx files"

        if [ "${NGENERATED}" -eq 0 ]; then
            echo "[entrypoint] WARNING: no output for variant ${VNAME}, skipping"
            continue
        fi

        VARIANT_REPORT="${REPORT_DIR}/${VNAME}"
        mkdir -p "${VARIANT_REPORT}" "${WORK_DIR}/${VNAME}"

        echo "[entrypoint] SSIM comparison for variant: ${VNAME} …"
        python3 /opt/scripts/compare_ssim.py \
            --original-dir="${MARK_DIR}" \
            --roundtrip-dir="${VDIR}" \
            --work-dir="${WORK_DIR}/${VNAME}" \
            --report="${VARIANT_REPORT}/index.html" \
            --threshold="${THRESHOLD}" \
            --dpi="${DPI}" \
            --workers="${WORKERS}" \
            || true
    done

    echo ""
    echo "=============================================="
    echo " Reports:"
    for VDIR in "${REPORT_DIR}"/*/; do
        [ -d "${VDIR}" ] || continue
        VNAME=$(basename "${VDIR}")
        echo "   ${VNAME}: replace-user-mark-batch/report/${VNAME}/index.html"
    done
    echo "=============================================="
    exit 0
fi

# ==========================================================================
# opc / docx: classic roundtrip pipeline
# ==========================================================================

echo "=============================================="
echo " ${LABEL} Visual Regression Test"
echo "=============================================="
echo " Layer:     ${LAYER}"
echo " Threshold: ${THRESHOLD}"
echo " DPI:       ${DPI}"
echo " Workers:   ${WORKERS}"
echo "=============================================="

NFILES=$(find "${ORIG_DIR}" -maxdepth 1 -iname '*.docx' | wc -l)
echo "[entrypoint] found ${NFILES} .docx files in corpus"

if [ "${NFILES}" -eq 0 ]; then
    echo "[entrypoint] ERROR: no .docx files found in ${ORIG_DIR}"
    echo "[entrypoint] Put your .docx files into visual-regtest/test-files/ and rebuild."
    exit 1
fi

# --------------------------------------------------------------------------
# Step 1: Run roundtrip on all .docx files.
# --------------------------------------------------------------------------
mkdir -p "${RT_DIR}"
echo "[entrypoint] running ${LABEL} roundtrip …"
"${ROUNDTRIP_BIN}" --input="${ORIG_DIR}" --output="${RT_DIR}" --workers="${WORKERS}" --layer="${LAYER}"

# --------------------------------------------------------------------------
# Step 2: SSIM comparison + report.
# --------------------------------------------------------------------------
echo "[entrypoint] running SSIM comparison …"
python3 /opt/scripts/compare_ssim.py \
    --original-dir="${ORIG_DIR}" \
    --roundtrip-dir="${RT_DIR}" \
    --work-dir="${WORK_DIR}" \
    --report="${REPORT_DIR}/index.html" \
    --threshold="${THRESHOLD}" \
    --dpi="${DPI}" \
    --workers="${WORKERS}" \
    || true  # don't fail the container; the report has the details

echo ""
echo "=============================================="
echo " Report: visual-regtest/report/index.html"
echo "=============================================="