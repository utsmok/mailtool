# Mailtool Performance Benchmarks

This directory contains performance benchmarks for the mailtool MCP server implementations.

## Purpose

The benchmarks compare the legacy MCP server implementation (v2.2, `mcp_server.py`) against the new MCP SDK v2 implementation (v2.3, `src/mailtool/mcp/server.py`) to ensure the migration maintains or improves performance.

## Benchmarks

### `performance_benchmark.py`

Comprehensive performance benchmarking script that tests:

1. **Tool Execution Speed**: Measures the time taken to perform common operations
   - List emails (10, 50, 100 items)
   - Get email details by EntryID
   - List calendar events (7 days)
   - List tasks (active and all)
   - Search emails with SQL filters

2. **Memory Usage**: Tracks memory allocation during operations
   - Memory before and after each benchmark
   - Memory delta for each operation

3. **Memory Leak Detection**: Performs repeated operations to detect memory leaks
   - Repeated list_emails operations (20 iterations)
   - Repeated get_email operations (20 iterations)
   - Repeated list_tasks operations (20 iterations)
   - Monitors memory growth over iterations

## Usage

### Prerequisites

- Windows with Outlook installed and running
- pywin32 dependency installed
- Python 3.13+

### Running Benchmarks

From the project root directory:

```bash
# Using uv (recommended)
uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark

# Or using Python directly (if pywin32 is installed)
python -m scripts.benchmarks.performance_benchmark
```

### Understanding Results

The benchmark outputs two sections:

#### 1. Performance Benchmark Results

Shows timing statistics for each operation:
- **Iterations**: Number of successful runs
- **Total time**: Cumulative time across all iterations
- **Avg time**: Average time per iteration
- **Min/Max time**: Fastest and slowest iterations
- **Median time**: Middle value (more robust than average)
- **Std dev**: Standard deviation (consistency measure)
- **Memory delta**: Memory change during the operation

#### 2. Memory Leak Analysis

Shows memory growth patterns during repeated operations:
- **Initial memory**: Memory usage at start
- **Final memory**: Memory usage after all iterations
- **Growth**: Absolute and percentage change
- **Warning indicator**: If growth exceeds 10%, potential leak detected

## Success Criteria

The MCP SDK v2 implementation should:

1. **Performance**: Within 20% of legacy implementation speed
   - Faster is better, but slight overhead is acceptable for the benefits of structured output and type safety

2. **Memory Usage**: Stable memory usage with no leaks
   - Memory growth should be < 10% over 20 iterations
   - Memory should be released after operations complete

3. **Large Folders**: Consistent performance regardless of folder size
   - List operations should scale linearly with item count
   - No significant performance degradation with 100+ items

## Troubleshooting

### "Cannot connect to Outlook" Error

**Cause**: Outlook is not running or not accessible

**Solution**:
1. Start Outlook on Windows
2. Ensure Outlook is fully loaded (wait 5-10 seconds)
3. Run benchmark again

### Import Errors

**Cause**: pywin32 not installed

**Solution**:
```bash
uv run --with pywin32 python -m scripts.benchmarks.performance_benchmark
```

### Inconsistent Results

**Cause**: Background processes or Outlook state affecting performance

**Solution**:
1. Close other applications
2. Ensure Outlook is idle (no syncing in progress)
3. Run benchmark multiple times and average results

## Adding New Benchmarks

To add a new benchmark:

1. Add a new test case in `run_all_benchmarks()` method
2. Follow the existing pattern:
   ```python
   results.append(
       self.benchmark_operation(
           "Descriptive name",
           self.bridge.method_name,
           arg1, arg2,  # Method arguments
       )
   )
   ```

3. For memory leak tests, add to `run_memory_leak_test()` method
4. Update this README with new test descriptions

## Notes

- Benchmarks use `time.perf_counter()` for high-resolution timing
- Memory tracking uses `tracemalloc` for accurate allocation measurement
- Warmup iterations ensure JIT compilation and caching don't skew results
- Results are statistical in nature - run multiple times for confidence
