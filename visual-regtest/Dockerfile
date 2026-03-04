# ===========================================================================
# Visual Regression Test – self-contained Docker image
#
# Contains: Go 1.25 · LibreOffice (headless) · poppler-utils · Python 3 +
#           scikit-image, Pillow, numpy
#
# Test corpus is baked into the image from visual-regtest/test-files/.
#
# A single roundtrip binary is built, selected by the --layer flag:
#   roundtrip --layer=opc   – OPC layer only  (opc.OpenFile → SaveToFile)
#   roundtrip --layer=docx  – full docx layer (docx.OpenFile → doc.SaveFile)
#
# Select which layer to test via the LAYER env variable (default: docx).
# ===========================================================================
FROM golang:1.25-bookworm

# ---------- system dependencies -------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        poppler-utils \
        python3 \
        python3-pip \
        python3-venv \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

# ---------- Python dependencies -------------------------------------------
RUN python3 -m pip install --break-system-packages --no-cache-dir \
        scikit-image \
        Pillow \
        numpy

# ---------- project source & build ----------------------------------------
WORKDIR /workspace
COPY . .
RUN go build -o /usr/local/bin/roundtrip             ./visual-regtest/roundtrip \
 && go build -o /usr/local/bin/replace-mark-batch   ./visual-regtest/replace-user-mark-batch

# ---------- test corpus (baked in) ----------------------------------------
COPY visual-regtest/test-files/ /corpus/

# ---------- replace-mark-batch input (baked in) ---------------------------
COPY visual-regtest/replace-user-mark-batch/in/ /mark-batch-in/

# ---------- scripts -------------------------------------------------------
COPY visual-regtest/scripts/ /opt/scripts/
RUN chmod +x /opt/scripts/entrypoint.sh

# ---------- runtime -------------------------------------------------------
ENV LAYER=docx \
    SSIM_THRESHOLD=0.98 \
    DPI=150 \
    WORKERS=4

ENTRYPOINT ["/opt/scripts/entrypoint.sh"]