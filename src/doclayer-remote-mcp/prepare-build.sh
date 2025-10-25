#!/bin/bash
# Prepare Docker build context by copying python wrapper

echo "Copying python-wrapper to doclayer_python_local..."
rm -rf doclayer_python_local
cp -r ../../python-wrapper doclayer_python_local

echo "Verifying DLLs are present..."
if [ -f "doclayer_python_local/doclayer_python/bin/DocLayer.Core.dll" ]; then
    echo "✓ DocLayer.Core.dll found"
else
    echo "✗ ERROR: DocLayer.Core.dll not found! Make sure C# project is built first."
    echo "Run: cd ../../src/DocLayer.Core/DocLayer.Core && dotnet build"
    exit 1
fi

echo "Build context prepared. You can now run:"
echo "  docker build -t doclayer-mcp ."
echo "Or deploy to Render"
