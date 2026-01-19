"""
Lifespan Management for MCP Server

This module will contain the Outlook bridge lifecycle management.
It provides async context manager for:
- Creating and warming up OutlookBridge instance
- Releasing COM objects and forcing garbage collection on shutdown
"""

# TODO: Implement lifespan management in US-002
