# Expected Benchmark Results

This document contains expected benchmark results for the MCP SDK v2 implementation.

## Important Note

**These benchmarks require Windows with Outlook running and pywin32 installed.**

The benchmark script (`performance_benchmark.py`) cannot run in:
- WSL2 environments (without Windows Outlook access)
- CI/CD environments (without Outlook)
- Non-Windows platforms

To run these benchmarks, execute on Windows with Outlook running:

```bash
uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
```

## Expected Output Format

When run successfully, the benchmark produces two sections:

### 1. Performance Benchmark Results

```
============================================================
PERFORMANCE BENCHMARK RESULTS
============================================================

List 10 emails from Inbox:
  Iterations: 10
  Total time: 1.234s
  Avg time: 123.45ms
  Min time: 115.20ms
  Max time: 145.80ms
  Median time: 120.30ms
  Std dev: 10.25ms
  Memory delta: 12.34KB

List 50 emails from Inbox:
  Iterations: 10
  Total time: 2.456s
  Avg time: 245.60ms
  Min time: 230.10ms
  Max time: 278.90ms
  Median time: 240.50ms
  Std dev: 15.80ms
  Memory delta: 45.67KB

... (additional benchmark results)
```

### 2. Memory Leak Analysis

```
============================================================
MEMORY LEAK ANALYSIS
============================================================

list_emails:
  Initial memory: 1024.00KB
  Final memory: 1056.00KB
  Growth: 32.00KB (3.1%)
  OK: Memory usage stable

get_email:
  Initial memory: 1056.00KB
  Final memory: 1072.00KB
  Growth: 16.00KB (1.5%)
  OK: Memory usage stable

list_tasks:
  Initial memory: 1024.00KB
  Final memory: 1040.00KB
  Growth: 16.00KB (1.6%)
  OK: Memory usage stable
```

## Success Criteria

The MCP SDK v2 implementation should meet these criteria:

### Performance
- **Avg time**: Within 20% of legacy implementation
  - Example: If legacy takes 100ms, SDK v2 should take < 120ms
  - Slight overhead is acceptable for structured output and type safety benefits

### Memory Stability
- **Memory growth**: < 10% over 20 iterations
  - Indicates proper COM object cleanup
  - No memory leaks from unreleased references

### Large Folder Handling
- **List 100 emails**: Should complete without significant degradation
  - Performance should scale linearly with item count
  - No exponential slowdown with larger folders

## Legacy vs SDK v2 Comparison

### Legacy Implementation (v2.2)
- Direct JSON-RPC implementation
- Manual JSON serialization/deserialization
- No type safety (string-based returns)
- Manual error handling

### SDK v2 Implementation (v2.3)
- Official MCP Python SDK v2
- FastMCP framework with decorators
- Pydantic models for structured output
- Type-safe return values
- Built-in error handling with custom exceptions
- Async lifespan management

### Expected Performance Differences

The SDK v2 implementation may have slight overhead (5-15%) due to:
1. **Pydantic validation**: Model serialization adds ~1-5ms per call
2. **FastMCP framework**: Decorator and routing overhead (~1-3ms)
3. **Type checking**: Runtime type validation (~1-2ms)

However, this overhead is acceptable because:
- **Benefits**: Structured output, type safety, better error messages
- **LLM understanding**: Pydantic Field() descriptions improve AI comprehension
- **Maintainability**: Declarative tools are easier to extend and debug
- **Safety**: Compile-time and runtime type checking prevent bugs

## Manual Testing Instructions

To run benchmarks manually:

1. **Start Outlook** on Windows
2. **Wait** for Outlook to fully load (5-10 seconds)
3. **Open terminal** in project root
4. **Run benchmark**:
   ```bash
   uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
   ```
5. **Review results** for:
   - Avg time per operation
   - Memory growth percentage
   - Memory leak warnings

## CI/CD Limitations

These benchmarks are **NOT** run in CI/CD because:
- GitHub Actions runners don't have Outlook
- Cannot test COM bridge functionality without Outlook
- Memory testing requires actual COM objects

Alternative CI testing:
- Unit tests use mock bridge (tests/mcp/)
- Integration tests simulate operations with mocks
- Manual benchmarks run locally before releases

## Documenting Results

After running benchmarks, update this file with actual results:

```markdown
## Actual Results (YYYY-MM-DD)

### Performance
- List 10 emails: 120ms avg
- List 50 emails: 240ms avg
- Get email details: 85ms avg
- ... (etc)

### Memory Stability
- list_emails: 3.1% growth (PASS)
- get_email: 1.5% growth (PASS)
- list_tasks: 1.6% growth (PASS)

### Conclusion
SDK v2 implementation meets all performance criteria.
```
