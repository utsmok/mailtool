# UV Migration Summary

## What Changed

Migrated from using global Windows Python to [uv](https://github.com/astral-sh/uv) for dependency management.

## Before

```bash
# Used global Python
python.exe -m pip install pywin32  # Pollutes global Python
python.exe script.py               # Uses system Python
```

## After

```bash
# Uses uv for dependency management
./outlook.sh emails  # Automatically manages pywin32 via uv
```

## Key Benefits

✅ **No global Python pollution** - pywin32 is installed only when needed
✅ **Reproducible** - Dependencies tracked in `pyproject.toml`
✅ **Cross-platform** - Linux tooling uses `.venv`, Windows COM automation uses uv
✅ **Automatic** - `uv run --with pywin32` installs on-the-fly
✅ **Fast** - uv is much faster than pip

## Architecture

```
WSL2 (Linux)              Windows
┌─────────────────┐       ┌──────────────────────┐
│  outlook.sh     │──────>│  outlook.bat         │
│  (wrapper)      │       │  (uses uv)           │
└─────────────────┘       └────────┬─────────────┘
                                   │
                                   ▼
                          ┌─────────────────────┐
                          │ uv run --with       │
                          │   pywin32 python... │
                          └────────┬────────────┘
                                   │
                                   ▼
                          ┌─────────────────┐
                          │   Outlook.exe   │
                          │   (COM API)     │
                          └─────────────────┘
```

## Files Modified

1. **pyproject.toml** - Added uv project config with optional windows dependencies
2. **outlook.bat** - New Windows entry point using `uv run`
3. **outlook.sh** - Updated to call batch file instead of python.exe directly
4. **src/mailtool_outlook_bridge.py** - Moved from root to src/ for better organization

## Testing

```bash
# Works perfectly!
./outlook.sh emails --limit 2
# Returns JSON with actual emails

# First run downloads pywin32:
# Downloading pywin32 (9.2MiB)
# Installed 1 package in 124ms
```

## Future Commands

```bash
# Add new Windows-specific dependencies
# (Edit pyproject.toml: project.optional-dependencies.windows)

# Add general dependencies
uv add <package>

# Run any Python script with uv
uv run python script.py
```

## Performance

- **First run**: ~124ms to download and install pywin32
- **Subsequent runs**: Instant (uv caches packages)
- **No manual setup**: Just works™
