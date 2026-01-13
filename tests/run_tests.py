"""
Test Runner for Axe Annotate
=============================
Consolidates all test scripts for easy execution.

Usage:
    python run_tests.py             # Run all tests
    python run_tests.py debug       # Run debug_excel.py
    python run_tests.py stress      # Run stress_test_excel.py
    python run_tests.py queue       # Run test_queue.py
    python run_tests.py tabswitch   # Run diagnose_tab_switch.py
    python run_tests.py connection  # Run verify_connection.py
    
All tests run in non-interactive mode (--auto flag) by default.
"""

import subprocess
import sys
import os

# Add parent directory to path to find our modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

TESTS = {
    "debug": ("debug_excel.py", "Basic Excel connection and annotation test"),
    "stress": ("stress_test_excel.py", "Multi-iteration reliability test"),
    "queue": ("test_queue.py", "Worker queue pattern test"),
    "tabswitch": ("diagnose_tab_switch.py", "Tab/workbook switching test"),
    "connection": ("verify_connection.py", "Quick connection verification"),
}


def run_test(test_file, auto=True):
    """Run a test file with optional --auto flag."""
    test_path = os.path.join(os.path.dirname(__file__), test_file)
    cmd = [sys.executable, test_path]
    if auto and test_file != "verify_connection.py":
        cmd.append("--auto")
    
    print(f"\n{'='*60}")
    print(f"Running: {test_file}")
    print('='*60 + "\n")
    
    result = subprocess.run(cmd)
    return result.returncode == 0


def main():
    if len(sys.argv) > 1:
        test_name = sys.argv[1].lower()
        if test_name in TESTS:
            test_file, desc = TESTS[test_name]
            print(f"Running: {desc}")
            success = run_test(test_file)
            sys.exit(0 if success else 1)
        elif test_name == "all":
            # Run all tests
            pass
        else:
            print(f"Unknown test: {test_name}")
            print(f"Available: {', '.join(TESTS.keys())}, all")
            sys.exit(1)
    
    # Run all tests
    print("Running all tests...\n")
    results = {}
    
    for name, (test_file, desc) in TESTS.items():
        print(f"\n[{name}] {desc}")
        success = run_test(test_file)
        results[name] = success
    
    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    for name, success in results.items():
        status = "PASS" if success else "FAIL"
        print(f"  {name}: {status}")
    
    passed = sum(1 for s in results.values() if s)
    total = len(results)
    print(f"\nTotal: {passed}/{total} passed")
    
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
