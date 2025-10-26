# Accept base image as build argument for CI/CD flexibility
ARG BASE_IMAGE=openenv-base:latest
FROM ${BASE_IMAGE}

# Install LibreOffice and required dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-java-common \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-draw \
    libreoffice-math \
    libreoffice-base \
    libreoffice-l10n-en-us \
    libreoffice-help-en-us \
    python3-uno \
    && rm -rf /var/lib/apt/lists/*
RUN apt-get update && apt-get install -y tini && rm -rf /var/lib/apt/lists/*

ENTRYPOINT ["/usr/bin/tini", "--"]
# Install additional Python packages for LibreOffice integration
# RUN pip install --no-cache-dir \
#     uno-python \
#     unohelper


# ✅ Remove this block — these are already installed by python3-uno
# RUN pip install --no-cache-dir \
#     uno-python \
#     unohelper

# (Optional) Install FastAPI stack if not in openenv-base
RUN pip install --no-cache-dir fastapi uvicorn pydantic  
# Set environment variables for LibreOffice
ENV DISPLAY=:99
ENV LIBREOFFICE_HEADLESS=1

# Copy environment code
COPY src/core/ /app/src/core/
COPY src/envs/libreoffice_env/ /app/src/envs/libreoffice_env/

# Create a directory for temporary files
# Create a directory for temporary files
RUN mkdir -p /tmp/libreoffice_env && chmod 777 /tmp/libreoffice_env

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# Prepare LibreOffice runtime directory
RUN mkdir -p /var/run/libreoffice && chmod 777 /var/run/libreoffice

# The LibreOffice service must listen on a known socket for UNO to connect
ENV UNO_SOCKET_PORT=2002
ENV PYTHONPATH="/usr/lib/python3/dist-packages:${PYTHONPATH}"

# ✅ Final combined CMD — run both LibreOffice and FastAPI in the same process
# Start both LibreOffice headless service and FastAPI app
CMD ["bash", "-c", "trap 'pkill -TERM -f soffice.bin; pkill -TERM -f uvicorn; exit 0' SIGTERM SIGINT; soffice --headless --accept='socket,host=0,port=2002;urp;StarOffice.ServiceManager' --nologo --nodefault --nofirststartwizard & sleep 3 && uvicorn envs.libreoffice_env.server.app:app --host 0.0.0.0 --port 8000 & wait"]
