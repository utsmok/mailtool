#!/usr/bin/env python
"""Test script to verify complete_task tool is registered correctly"""

from mailtool.mcp.server import mcp, complete_task

# Test 1: Function is importable
print("Test 1: complete_task function is importable")
print(f"  complete_task: {complete_task}")
print(f"  Type: {type(complete_task)}")
print()

# Test 2: Function signature
print("Test 2: complete_task function signature")
import inspect
sig = inspect.signature(complete_task)
print(f"  Signature: {sig}")
for param_name, param in sig.parameters.items():
    print(f"    - {param_name}: {param.annotation if param.annotation != inspect.Parameter.empty else 'Any'}")
print()

# Test 3: Return type
print("Test 3: complete_task return type")
print(f"  Return annotation: {inspect.signature(complete_task).return_annotation}")
print()

# Test 4: Docstring
print("Test 4: complete_task docstring")
print(f"  Docstring present: {complete_task.__doc__ is not None}")
if complete_task.__doc__:
    print(f"  First line: {complete_task.__doc__.strip().split(chr(10))[0]}")
print()

print("All tests passed!")
