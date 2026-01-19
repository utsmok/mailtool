#!/usr/bin/env python3
"""Test script for lifespan management (structure verification only)

This script verifies that the lifespan code structure is correct.
Full integration test requires Outlook running on Windows.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))


def test_lifespan_structure():
    """Test that lifespan structure is correct (without actual COM calls)"""
    from mailtool.mcp.lifespan import (
        OutlookContext,
        _create_bridge,
        _warmup_bridge,
        outlook_lifespan,
    )

    print("Testing lifespan structure...")

    # Test that OutlookContext is a dataclass
    assert hasattr(OutlookContext, "__dataclass_fields__"), (
        "OutlookContext should be a dataclass"
    )
    print("[OK] OutlookContext is a dataclass")

    # Test that outlook_lifespan is decorated with asynccontextmanager
    # The decorator wraps it, so check for __aenter__ and __aexit__ on the result
    import contextlib

    assert isinstance(
        outlook_lifespan, contextlib.AbstractAsyncContextManager
    ) or hasattr(outlook_lifespan, "__wrapped__"), (
        "outlook_lifespan should be an async context manager"
    )
    print("[OK] outlook_lifespan is an async context manager")

    # Test that helper functions exist and are callable
    assert callable(_create_bridge), "_create_bridge should be callable"
    assert callable(_warmup_bridge), "_warmup_bridge should be callable"
    print("[OK] Helper functions are defined")

    # Test function signatures

    # Verify outlook_lifespan has the asynccontextmanager decorator
    assert hasattr(outlook_lifespan, "__wrapped__"), (
        "outlook_lifespan should be decorated"
    )
    print("[OK] outlook_lifespan has asynccontextmanager decorator")

    print("\nAll structure tests passed!")
    print("\nNote: Full integration test requires Outlook running on Windows.")


def test_imports():
    """Test that all imports work correctly"""
    print("\nTesting imports...")
    from mailtool.mcp import lifespan

    assert hasattr(lifespan, "OutlookContext")
    assert hasattr(lifespan, "outlook_lifespan")
    print("[OK] All imports successful")


if __name__ == "__main__":
    test_imports()
    test_lifespan_structure()
    print("\n" + "=" * 70)
    print("SUCCESS: Lifespan module structure is correct!")
    print("=" * 70)
