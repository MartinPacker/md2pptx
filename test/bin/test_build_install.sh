#!/bin/bash

# Test script for md2pptx project
# Uses uv tool to build, install, and test the project

# Navigate to project root directory
cd "$(dirname "$(dirname "$(dirname "$0")")")"

echo "=== Testing md2pptx project ==="
echo ""

# Check if uv is available
if ! command -v uv &> /dev/null; then
    echo "Error: uv command not found!"
    echo "Please install uv first: https://docs.astral.sh/uv/getting-started/"
    exit 1
fi
echo "uv is available."
echo ""

# Step 1: Build the project
echo "1. Building the project..."
uv build
if [ $? -ne 0 ]; then
    echo "Build failed!"
    exit 1
fi
echo "Build completed successfully!"
echo ""

# Step 2: Install the project
echo "2. Installing the project..."
uv pip install --force-reinstall dist/*.whl
if [ $? -ne 0 ]; then
    echo "Installation failed!"
    exit 1
fi
echo "Installation completed successfully!"
echo ""

# Step 3: Test the installation
echo "3. Testing the installation..."

# Check if md2pptx command is available
if ! command -v md2pptx &> /dev/null; then
    echo "Error: md2pptx command not found!"
    exit 1
fi
echo "md2pptx command is available."

# Test with a simple markdown file
cat > test_input.md << 'EOF'
template: Martin Template.pptx

# Test Presentation

## Section 1

### Slide 1

* Item 1
* Item 2
* Item 3
EOF

echo "Created test input file."

# Run md2pptx with the test input
md2pptx test_output.pptx < test_input.md
if [ $? -ne 0 ]; then
    echo "Error: md2pptx command failed!"
    exit 1
fi
echo "md2pptx command executed successfully."

# Check if output file was created and has non-zero size
if [ -f test_output.pptx ]; then
    if [ -s test_output.pptx ]; then
        echo "Output file test_output.pptx was created successfully with non-zero size!"
    else
        echo "Error: Output file was created but has zero size!"
        exit 1
    fi
else
    echo "Error: Output file was not created!"
    exit 1
fi
echo ""
echo "=== All tests passed! ==="

# Clean up
rm -f test_input.md test_output.pptx

# Clean up build artifacts
rm -rf dist build *.egg-info

echo "Cleanup completed."
