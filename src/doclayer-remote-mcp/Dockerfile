FROM python:3.11-slim

# Install .NET 8.0 runtime
RUN apt-get update && \
    apt-get install -y wget && \
    wget https://dot.net/v1/dotnet-install.sh -O dotnet-install.sh && \
    chmod +x dotnet-install.sh && \
    ./dotnet-install.sh --channel 8.0 --runtime dotnet --install-dir /usr/share/dotnet && \
    ln -s /usr/share/dotnet/dotnet /usr/bin/dotnet && \
    rm dotnet-install.sh && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Install dependencies
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Clone and install DocLayer Python wrapper from GitHub
RUN pip install --no-cache-dir pythonnet>=3.0.0 && \
    pip install git+https://github.com/shray-ya-sing/doclayer.git#subdirectory=python-wrapper

# Copy remote MCP server code
COPY server.py /app/

# Expose port
EXPOSE 8000

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Run server
CMD ["python", "server.py"]
