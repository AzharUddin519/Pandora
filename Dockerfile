FROM mcr.microsoft.com/vscode/devcontainers/base:ubuntu

# Install PowerShell
RUN apt-get update \
    && apt-get install -y curl gnupg apt-transport-https \
    && curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - \
    && curl https://packages.microsoft.com/config/ubuntu/20.04/prod.list > /etc/apt/sources.list.d/microsoft.list \
    && apt-get update \
    && apt-get install -y powershell

# Clean up
RUN apt-get clean && rm -rf /var/lib/apt/lists/*

# Set PowerShell as the default shell
ENV SHELL /usr/bin/pwsh
