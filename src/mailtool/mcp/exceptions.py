"""Custom exception classes for MCP server.

This module provides domain-specific exception classes that extend McpError
for better error handling and debugging in the MCP server.
"""

from mcp import McpError
from mcp.shared.exceptions import ErrorData

# Error codes for custom exceptions
ERROR_CODE_NOT_FOUND = -32602  # Item not found
ERROR_CODE_COM_ERROR = -32603  # COM/bridge error
ERROR_CODE_VALIDATION_ERROR = -32604  # Validation error


class OutlookNotFoundError(McpError):
    """Exception raised when an Outlook item is not found.

    This exception is raised when attempting to access an email, appointment,
    or task that doesn't exist or cannot be found via its EntryID.

    Attributes:
        message: Error message describing what was not found
        entry_id: The EntryID that was not found (optional)
    """

    def __init__(self, message: str, entry_id: str | None = None) -> None:
        """Initialize the exception.

        Args:
            message: Error message describing what was not found
            entry_id: The EntryID that was not found (optional)
        """
        if entry_id:
            full_message = f"{message} (EntryID: {entry_id})"
        else:
            full_message = message

        error_data = ErrorData(
            code=ERROR_CODE_NOT_FOUND,
            message=full_message,
            data={"entry_id": entry_id} if entry_id else None,
        )
        super().__init__(error_data)
        self.entry_id = entry_id


class OutlookComError(McpError):
    """Exception raised when Outlook COM operation fails.

    This exception is raised when there's a failure in the underlying
    COM bridge or Outlook application. This includes:
    - Bridge not initialized
    - COM object access failures
    - Outlook application errors
    - Thread pool executor failures

    Attributes:
        message: Error message describing the COM failure
        details: Additional error details (optional)
    """

    def __init__(self, message: str, details: str | None = None) -> None:
        """Initialize the exception.

        Args:
            message: Error message describing the COM failure
            details: Additional error details (optional)
        """
        if details:
            full_message = f"{message}: {details}"
        else:
            full_message = message

        error_data = ErrorData(
            code=ERROR_CODE_COM_ERROR,
            message=full_message,
            data={"details": details} if details else None,
        )
        super().__init__(error_data)
        self.details = details


class OutlookValidationError(McpError):
    """Exception raised when input validation fails.

    This exception is raised when user input doesn't meet requirements,
    such as:
    - Invalid date formats
    - Invalid parameter values
    - Missing required fields
    - Invalid email addresses

    Attributes:
        message: Error message describing the validation failure
        field: The field that failed validation (optional)
    """

    def __init__(self, message: str, field: str | None = None) -> None:
        """Initialize the exception.

        Args:
            message: Error message describing the validation failure
            field: The field that failed validation (optional)
        """
        if field:
            full_message = f"Validation failed for '{field}': {message}"
        else:
            full_message = f"Validation failed: {message}"

        error_data = ErrorData(
            code=ERROR_CODE_VALIDATION_ERROR,
            message=full_message,
            data={"field": field} if field else None,
        )
        super().__init__(error_data)
        self.field = field
