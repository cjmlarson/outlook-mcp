#!/usr/bin/env python3
"""
Test script to verify calendar search/filter fixes
"""

import subprocess
import json
import sys

def run_command(cmd):
    """Run a command and return the output"""
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    return result.stdout, result.stderr, result.returncode

def test_calendar_filter():
    """Test calendar filtering"""
    print("Testing calendar filter...")
    cmd = 'py python/outlook_filter.py "connor.larson@outlook.com/Calendar" --days 30 --type event'
    stdout, stderr, code = run_command(cmd)
    
    if code != 0:
        print(f"  [FAIL] Filter failed: {stderr}")
        return False
    
    try:
        data = json.loads(stdout)
        if data['total'] > 0:
            print(f"  [PASS] Found {data['total']} calendar events")
            # Check if we found known events
            subjects = [r['subject'] for r in data['results']]
            if any('CoffeeRun' in s for s in subjects):
                print(f"  [PASS] Found CoffeeRun event")
            return True
        else:
            print(f"  [WARN] No calendar events found")
            return False
    except json.JSONDecodeError:
        print(f"  [FAIL] Invalid JSON output")
        return False

def test_calendar_search():
    """Test calendar search"""
    print("\nTesting calendar search...")
    cmd = 'py python/outlook_search.py "Monday" "connor.larson@outlook.com/Calendar"'
    stdout, stderr, code = run_command(cmd)
    
    if code != 0:
        print(f"  [FAIL] Search failed: {stderr}")
        return False
    
    # Check for the error that was happening before
    if "ReceivedTime" in stderr and "unknown" in stderr:
        print(f"  [FAIL] Still using ReceivedTime for calendar: {stderr}")
        return False
    
    try:
        data = json.loads(stdout)
        if data['pagination']['total'] > 0:
            print(f"  [PASS] Found {data['pagination']['total']} items matching 'Monday'")
            return True
        else:
            print(f"  [WARN] No items found for 'Monday'")
            return False
    except json.JSONDecodeError:
        print(f"  [FAIL] Invalid JSON output")
        return False

def test_email_regression():
    """Ensure email functionality still works"""
    print("\nTesting email functionality (regression check)...")
    
    # Test email filter
    cmd = 'py python/outlook_filter.py "connor.larson@outlook.com/Inbox" --days 30'
    stdout, stderr, code = run_command(cmd)
    
    if code != 0:
        print(f"  [FAIL] Email filter failed: {stderr}")
        return False
    
    try:
        data = json.loads(stdout)
        print(f"  [PASS] Email filter works: {data['total']} emails in Inbox")
    except:
        print(f"  [FAIL] Email filter output invalid")
        return False
    
    # Test email search
    cmd = 'py python/outlook_search.py "meeting" "connor.larson@outlook.com/Sent Items"'
    stdout, stderr, code = run_command(cmd)
    
    if code != 0:
        print(f"  [FAIL] Email search failed: {stderr}")
        return False
    
    try:
        data = json.loads(stdout)
        print(f"  [PASS] Email search works: {data['pagination']['total']} results")
        return True
    except:
        print(f"  [FAIL] Email search output invalid")
        return False

def main():
    print("=" * 50)
    print("CALENDAR FIX VERIFICATION TEST SUITE")
    print("=" * 50)
    
    tests_passed = 0
    tests_total = 3
    
    if test_calendar_filter():
        tests_passed += 1
    
    if test_calendar_search():
        tests_passed += 1
    
    if test_email_regression():
        tests_passed += 1
    
    print("\n" + "=" * 50)
    print(f"RESULTS: {tests_passed}/{tests_total} tests passed")
    
    if tests_passed == tests_total:
        print("[PASS] All tests passed! Calendar fix is working correctly.")
        return 0
    else:
        print("[FAIL] Some tests failed. Please review the output above.")
        return 1

if __name__ == "__main__":
    sys.exit(main())