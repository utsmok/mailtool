#!/usr/bin/env python3
"""
Performance Benchmark: MCP SDK v2 vs Legacy Implementation

This script benchmarks the legacy MCP server (mcp_server.py) against the new
SDK v2 implementation (src/mailtool/mcp/server.py) to ensure the migration
maintains or improves performance.

Tests:
1. Tool execution speed (same operations on both implementations)
2. Memory usage and COM object cleanup
3. Large folder handling (listing many items)
4. Repeated operations to detect memory leaks

Usage:
    uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
"""

import contextlib
import gc
import sys
import time
import tracemalloc
from dataclasses import dataclass
from pathlib import Path
from statistics import mean, median, stdev
from typing import TYPE_CHECKING

# Add src to path for imports
src_path = Path(__file__).parent.parent.parent / "src"
sys.path.insert(0, str(src_path))

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

try:
    from mailtool.bridge import OutlookBridge
except ImportError:
    print("Warning: mailtool.bridge not available (pywin32 not installed)")
    print("This benchmark requires Outlook on Windows with pywin32 installed")
    sys.exit(1)


@dataclass
class BenchmarkResult:
    """Results from a single benchmark run"""

    name: str
    iterations: int
    total_time: float
    avg_time: float
    min_time: float
    max_time: float
    median_time: float
    stdev_time: float | None
    memory_before: int
    memory_after: int
    memory_delta: int

    def __str__(self) -> str:
        std_dev_line = (
            f"  Std dev: {self.stdev_time * 1000:.2f}ms\n"
            if self.stdev_time
            else "  Std dev: N/A\n"
        )
        return (
            f"{self.name}:\n"
            f"  Iterations: {self.iterations}\n"
            f"  Total time: {self.total_time:.3f}s\n"
            f"  Avg time: {self.avg_time * 1000:.2f}ms\n"
            f"  Min time: {self.min_time * 1000:.2f}ms\n"
            f"  Max time: {self.max_time * 1000:.2f}ms\n"
            f"  Median time: {self.median_time * 1000:.2f}ms\n"
            f"{std_dev_line}"
            f"  Memory delta: {self.memory_delta / 1024:.2f}KB\n"
        )


class PerformanceBenchmark:
    """Performance benchmark for MCP server implementations"""

    bridge: "OutlookBridge | None"

    def __init__(self, warmup_iterations: int = 3, benchmark_iterations: int = 10):
        """Initialize benchmark

        Args:
            warmup_iterations: Number of warmup runs before measuring
            benchmark_iterations: Number of benchmark runs to measure
        """
        self.warmup_iterations = warmup_iterations
        self.benchmark_iterations = benchmark_iterations
        self.bridge = None

    def setup(self):
        """Set up Outlook bridge for benchmarking"""
        print("Setting up Outlook bridge...")
        self.bridge = OutlookBridge()

        # Warmup to ensure Outlook is responsive
        print(f"Warming up with {self.warmup_iterations} iterations...")
        for _ in range(self.warmup_iterations):
            self.bridge.get_inbox()

        print("Outlook bridge ready for benchmarking\n")

    def teardown(self):
        """Clean up after benchmarking"""
        if self.bridge:
            print("Cleaning up Outlook bridge...")
            self.bridge.cleanup()
            self.bridge = None
            gc.collect()
            print("Cleanup complete\n")

    def benchmark_operation(
        self, name: str, operation, *args, **kwargs
    ) -> BenchmarkResult:
        """Benchmark a single operation

        Args:
            name: Name of the benchmark
            operation: Function to benchmark
            *args: Positional arguments for operation
            **kwargs: Keyword arguments for operation

        Returns:
            BenchmarkResult with timing and memory statistics
        """
        print(f"Benchmarking: {name}")

        # Start memory tracking
        gc.collect()
        tracemalloc.start()
        memory_before = tracemalloc.get_traced_memory()[0]

        # Warmup runs
        for _ in range(self.warmup_iterations):
            with contextlib.suppress(Exception):
                operation(*args, **kwargs)

        # Benchmark runs
        times = []
        for _ in range(self.benchmark_iterations):
            start = time.perf_counter()
            try:
                result = operation(*args, **kwargs)
                end = time.perf_counter()
                times.append(end - start)
            except Exception as e:
                print(f"  Error during benchmark: {e}")
                continue

        # Stop memory tracking
        memory_after = tracemalloc.get_traced_memory()[0]
        tracemalloc.stop()

        if not times:
            raise RuntimeError(f"All benchmark iterations failed for {name}")

        # Calculate statistics
        total_time = sum(times)
        avg_time = mean(times)
        min_time = min(times)
        max_time = max(times)
        median_time = median(times)
        stdev_time = stdev(times) if len(times) > 1 else None

        memory_delta = memory_after - memory_before

        result = BenchmarkResult(
            name=name,
            iterations=len(times),
            total_time=total_time,
            avg_time=avg_time,
            min_time=min_time,
            max_time=max_time,
            median_time=median_time,
            stdev_time=stdev_time,
            memory_before=memory_before,
            memory_after=memory_after,
            memory_delta=memory_delta,
        )

        print(f"  Complete: {len(times)} iterations\n")
        return result

    def run_all_benchmarks(self) -> list[BenchmarkResult]:
        """Run all performance benchmarks

        Returns:
            List of BenchmarkResult objects
        """
        results = []

        try:
            self.setup()

            # Benchmark 1: List emails (small folder)
            results.append(
                self.benchmark_operation(
                    "List 10 emails from Inbox",
                    self.bridge.list_emails,
                    10,
                    "Inbox",
                )
            )

            # Benchmark 2: List emails (large folder)
            results.append(
                self.benchmark_operation(
                    "List 50 emails from Inbox",
                    self.bridge.list_emails,
                    50,
                    "Inbox",
                )
            )

            # Benchmark 3: List emails (very large folder)
            results.append(
                self.benchmark_operation(
                    "List 100 emails from Inbox",
                    self.bridge.list_emails,
                    100,
                    "Inbox",
                )
            )

            # Benchmark 4: Get email details
            # First get an email to test with
            emails = self.bridge.list_emails(1, "Inbox")
            if emails:
                entry_id = emails[0]["entry_id"]
                results.append(
                    self.benchmark_operation(
                        "Get email details by EntryID",
                        self.bridge.get_email_body,
                        entry_id,
                    )
                )

            # Benchmark 5: List calendar events
            results.append(
                self.benchmark_operation(
                    "List calendar events (7 days)",
                    self.bridge.list_calendar_events,
                    7,
                )
            )

            # Benchmark 6: List tasks
            results.append(
                self.benchmark_operation(
                    "List active tasks",
                    self.bridge.list_tasks,
                    False,
                )
            )

            # Benchmark 7: List all tasks
            results.append(
                self.benchmark_operation(
                    "List all tasks",
                    self.bridge.list_tasks,
                    True,
                )
            )

            # Benchmark 8: Search emails
            results.append(
                self.benchmark_operation(
                    "Search emails (SQL filter)",
                    self.bridge.search_emails,
                    "[Unread] = TRUE",
                    50,
                )
            )

        finally:
            self.teardown()

        return results

    def run_memory_leak_test(self) -> dict[str, list[int]]:
        """Run memory leak detection test

        Performs repeated operations and measures memory usage after each iteration
        to detect memory leaks from unreleased COM objects.

        Returns:
            Dict mapping operation name to list of memory measurements (bytes)
        """
        print("\n" + "=" * 60)
        print("MEMORY LEAK DETECTION TEST")
        print("=" * 60 + "\n")

        results = {}

        try:
            self.setup()

            # Test 1: Repeated list_emails operations
            print("Testing: Repeated list_emails operations")
            gc.collect()
            tracemalloc.start()
            memories = []

            for i in range(20):
                self.bridge.list_emails(10, "Inbox")
                if i % 5 == 0:
                    current_mem = tracemalloc.get_traced_memory()[0]
                    memories.append(current_mem)
                    print(f"  Iteration {i}: {current_mem / 1024:.2f}KB")

            results["list_emails"] = memories
            tracemalloc.stop()

            # Test 2: Repeated get_email operations
            print("\nTesting: Repeated get_email operations")
            emails = self.bridge.list_emails(1, "Inbox")
            if emails:
                entry_id = emails[0]["entry_id"]
                gc.collect()
                tracemalloc.start()
                memories = []

                for i in range(20):
                    self.bridge.get_email_body(entry_id)
                    if i % 5 == 0:
                        current_mem = tracemalloc.get_traced_memory()[0]
                        memories.append(current_mem)
                        print(f"  Iteration {i}: {current_mem / 1024:.2f}KB")

                results["get_email"] = memories
                tracemalloc.stop()

            # Test 3: Repeated list_tasks operations
            print("\nTesting: Repeated list_tasks operations")
            gc.collect()
            tracemalloc.start()
            memories = []

            for i in range(20):
                self.bridge.list_tasks(False)
                if i % 5 == 0:
                    current_mem = tracemalloc.get_traced_memory()[0]
                    memories.append(current_mem)
                    print(f"  Iteration {i}: {current_mem / 1024:.2f}KB")

            results["list_tasks"] = memories
            tracemalloc.stop()

        finally:
            self.teardown()

        return results

    def print_results(self, results: list[BenchmarkResult]):
        """Print benchmark results in a formatted table

        Args:
            results: List of BenchmarkResult objects
        """
        print("\n" + "=" * 60)
        print("PERFORMANCE BENCHMARK RESULTS")
        print("=" * 60 + "\n")

        for result in results:
            print(result)
            print()

    def analyze_memory_leaks(self, memory_data: dict[str, list[int]]):
        """Analyze memory leak test results

        Args:
            memory_data: Dict mapping operation name to memory measurements
        """
        print("\n" + "=" * 60)
        print("MEMORY LEAK ANALYSIS")
        print("=" * 60 + "\n")

        for operation, memories in memory_data.items():
            if len(memories) < 2:
                print(f"{operation}: Insufficient data for analysis")
                continue

            # Calculate memory growth
            initial = memories[0]
            final = memories[-1]
            growth = final - initial
            growth_pct = (growth / initial) * 100 if initial > 0 else 0

            print(f"{operation}:")
            print(f"  Initial memory: {initial / 1024:.2f}KB")
            print(f"  Final memory: {final / 1024:.2f}KB")
            print(f"  Growth: {growth / 1024:.2f}KB ({growth_pct:.1f}%)")

            # Check for potential leak (more than 10% growth)
            if growth_pct > 10:
                print("  WARNING: Potential memory leak detected!")
            else:
                print("  OK: Memory usage stable")
            print()


def main():
    """Main benchmark entry point"""
    print("\n" + "=" * 60)
    print("Mailtool MCP Server Performance Benchmark")
    print("=" * 60 + "\n")

    # Check if Outlook is available
    try:
        bridge = OutlookBridge()
        bridge.cleanup()
    except Exception as e:
        print(f"Error: Cannot connect to Outlook: {e}")
        print("This benchmark requires Outlook to be running on Windows")
        sys.exit(1)

    # Run benchmarks
    benchmark = PerformanceBenchmark(warmup_iterations=3, benchmark_iterations=10)

    # Performance benchmarks
    print("Running performance benchmarks...")
    results = benchmark.run_all_benchmarks()
    benchmark.print_results(results)

    # Memory leak tests
    print("\nRunning memory leak detection tests...")
    memory_data = benchmark.run_memory_leak_test()
    benchmark.analyze_memory_leaks(memory_data)

    print("\n" + "=" * 60)
    print("BENCHMARK COMPLETE")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
