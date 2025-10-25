#!/bin/bash
# Prepare Docker build context by copying python wrapper

echo "Copying python-wrapper to doclayer_python_local..."
rm -rf doclayer_python_local
cp -r ../../python-wrapper doclayer_python_local

echo "Build context prepared. You can now run:"
echo "  docker build -t doclayer-mcp ."
echo "Or deploy to Render"
