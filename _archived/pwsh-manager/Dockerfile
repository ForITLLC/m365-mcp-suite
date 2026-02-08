# Use official PnP PowerShell image - handles Linux compatibility correctly
# See: https://pnp.github.io/powershell/articles/docker.html
FROM m365pnp/powershell:nightly

# Install Python and additional PowerShell modules
# PnP image is Debian-based, use apt
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install additional PowerShell modules (PnP.PowerShell already included)
RUN pwsh -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted" && \
    pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force" && \
    pwsh -Command "Install-Module -Name Az.Accounts -Scope AllUsers -Force" && \
    pwsh -Command "Install-Module -Name MicrosoftTeams -Scope AllUsers -Force"

# Install .NET SDK and PAC CLI for Power Platform (PowerShell module doesn't work on Linux)
RUN curl -sSL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin --channel 8.0 --install-dir /usr/local/dotnet && \
    /usr/local/dotnet/dotnet tool install --global Microsoft.PowerApps.CLI.Tool --version 1.44.2
ENV DOTNET_ROOT=/usr/local/dotnet
ENV PATH="${PATH}:/usr/local/dotnet:/root/.dotnet/tools"

# Create app directory
WORKDIR /app

# Install Python dependencies (no venv needed in container)
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application code
COPY session_manager.py .

# Create data directory for token persistence
RUN mkdir -p /data/tokens

# Expose HTTP API port
EXPOSE 5100

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
    CMD curl -f http://localhost:5100/health || exit 1

# Override base image entrypoint and run the session manager
ENTRYPOINT []
CMD ["python3", "session_manager.py"]
