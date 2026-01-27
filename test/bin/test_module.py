#!/usr/bin/env python3

"""
Test script for md2pptx module
Tests the core functionality by importing the module directly
"""

import sys
import os
import tempfile
from io import StringIO

# Add the src directory to the path so we can import the module
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'src'))

try:
    from md2pptx import main
    print("✓ Successfully imported md2pptx module")
except ImportError as e:
    print(f"✗ Failed to import md2pptx module: {e}")
    sys.exit(1)

# Test with a simple markdown content
def test_module_functionality():
    print("\nTesting module functionality...")
    
    # Create test markdown content
    test_content = """
template: Martin Template.pptx

# Test Presentation

## Section 1

### Slide 1

* Item 1
* Item 2
* Item 3
"""
    
    # Create a temporary output file
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as temp_output:
        output_file = temp_output.name
    
    try:
        # Redirect stdin to our test content
        original_stdin = sys.stdin
        sys.stdin = StringIO(test_content)
        
        # Call main with the output file as argument
        sys.argv = ['md2pptx', output_file]
        main()
        
        print(f"✓ Successfully generated presentation: {output_file}")
        
        # Verify the output file exists and has content
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("✓ Output file exists and has content")
        else:
            print("✗ Output file is missing or empty")
            return False
        
    except Exception as e:
        print(f"✗ Error during module test: {e}")
        return False
    finally:
        # Restore original stdin
        sys.stdin = original_stdin
        
        # Clean up temporary file
        if os.path.exists(output_file):
            os.remove(output_file)
    
    return True

if __name__ == "__main__":
    print("=== Testing md2pptx module ===")
    
    if test_module_functionality():
        print("\n✅ All module tests passed!")
        sys.exit(0)
    else:
        print("\n❌ Some module tests failed!")
        sys.exit(1)
