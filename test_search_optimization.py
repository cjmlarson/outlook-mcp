#!/usr/bin/env python3
"""
Test script for validating outlook_search token optimization.
Tests the reduced token usage and ensures functionality is preserved.
"""

import sys
import json
import subprocess
from datetime import datetime, timedelta

def run_search(pattern, path, output_mode='list', offset=0):
    """Run outlook_search and return parsed output"""
    cmd = [
        'python', 
        'python/outlook_search.py', 
        pattern, 
        path,
        '--output-mode', output_mode,
        '--offset', str(offset)
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        return json.loads(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Error running search: {e.stderr}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        print(f"Output was: {result.stdout}")
        return None

def estimate_tokens(obj):
    """Rough estimate of token count for JSON object"""
    json_str = json.dumps(obj, default=str)
    # Rough approximation: ~4 characters per token
    return len(json_str) // 4

def test_pagination():
    """Test that pagination is now 10 items"""
    print("Testing pagination (should be 10 items max)...")
    
    # Search for common terms in Archive folder
    result = run_search("the", "connor.larson@outlook.com/Archive")
    
    if result:
        pagination = result.get('pagination', {})
        limit = pagination.get('limit')
        results_count = len(result.get('results', []))
        
        print(f"  Pagination limit: {limit} (expected: 10)")
        print(f"  Results returned: {results_count}")
        print(f"  Total available: {pagination.get('total', 0)}")
        
        if limit == 10:
            print("  ✓ Pagination correctly set to 10")
        else:
            print(f"  ✗ Pagination incorrect: {limit}")
        
        return limit == 10
    return False

def test_date_format():
    """Test compact date formatting"""
    print("\nTesting date format...")
    
    result = run_search("flight", "connor.larson@outlook.com/Archive")
    
    if result and result.get('results'):
        sample = result['results'][0]
        date_field = sample.get('received') or sample.get('date')
        
        print(f"  Sample date: {date_field}")
        
        # Check if it's in compact format (not ISO)
        if date_field and 'T' not in str(date_field) and len(str(date_field)) < 20:
            print("  ✓ Date in compact format")
            return True
        else:
            print("  ✗ Date not in compact format")
            return False
    
    print("  No results to test")
    return False

def test_field_optimization():
    """Test that unnecessary fields are removed"""
    print("\nTesting field optimization...")
    
    result = run_search("Swiss", "connor.larson@outlook.com/Archive")
    
    if result and result.get('results'):
        sample = result['results'][0]
        
        # Check for expected fields
        expected = ['entry_id', 'subject', 'sender', 'received']
        optional = ['has_attachments', 'is_read', 'matches']
        removed = ['size', 'importance', 'categories', 'path', 'from', 'date']
        
        print(f"  Fields present: {list(sample.keys())}")
        
        # Check removed fields are gone
        removed_found = [f for f in removed if f in sample]
        if removed_found:
            print(f"  ✗ Unexpected fields still present: {removed_found}")
            return False
        
        # Check required fields exist
        missing = [f for f in expected if f not in sample and f != 'received']
        if missing:
            print(f"  ✗ Required fields missing: {missing}")
            return False
        
        print("  ✓ Field optimization successful")
        return True
    
    print("  No results to test")
    return False

def test_token_comparison():
    """Compare token usage before and after optimization"""
    print("\nTesting token usage reduction...")
    
    # Search with list mode
    result_list = run_search("United", "connor.larson@outlook.com/Archive", 'list')
    
    if result_list:
        tokens_estimate = estimate_tokens(result_list)
        results_count = len(result_list.get('results', []))
        
        print(f"  Results: {results_count} items")
        print(f"  Estimated tokens: ~{tokens_estimate}")
        print(f"  Tokens per result: ~{tokens_estimate // max(results_count, 1)}")
        
        # Expected: ~80 tokens per result (vs old ~200)
        per_result = tokens_estimate // max(results_count, 1)
        if per_result < 120:  # Give some margin
            print("  ✓ Token usage reduced significantly")
            return True
        else:
            print("  ✗ Token usage still high")
            return False
    
    return False

def test_content_mode():
    """Test content mode with snippet optimization"""
    print("\nTesting content mode optimization...")
    
    result = run_search("ZRH", "connor.larson@outlook.com/Archive", 'content')
    
    if result and result.get('results'):
        sample = result['results'][0]
        
        if 'matches' in sample:
            match_count = len(sample['matches'])
            print(f"  Matches limited to: {match_count} (max 2 expected)")
            
            if match_count <= 2:
                print("  ✓ Content mode optimized")
                return True
            else:
                print("  ✗ Too many match snippets")
                return False
        else:
            print("  No matches found in content mode")
    
    return True

def test_entry_id_compatibility():
    """Test that entry_id is still compatible with outlook_read"""
    print("\nTesting entry_id compatibility...")
    
    result = run_search("Swiss", "connor.larson@outlook.com/Archive", 'list', 0)
    
    if result and result.get('results'):
        sample = result['results'][0]
        entry_id = sample.get('entry_id')
        
        if entry_id:
            print(f"  Entry ID present: {entry_id[:20]}...")
            
            # Test if we can read it
            cmd = ['python', 'python/outlook_read.py', entry_id]
            try:
                read_result = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
                if read_result.returncode == 0:
                    print("  ✓ Entry ID compatible with outlook_read")
                    return True
                else:
                    print(f"  ✗ outlook_read failed: {read_result.stderr}")
                    return False
            except Exception as e:
                print(f"  ✗ Error testing outlook_read: {e}")
                return False
    
    print("  No results to test")
    return False

def main():
    print("=" * 60)
    print("Outlook Search Token Optimization Test Suite")
    print("=" * 60)
    
    tests = [
        test_pagination,
        test_date_format,
        test_field_optimization,
        test_token_comparison,
        test_content_mode,
        test_entry_id_compatibility
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            if test():
                passed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"  ✗ Test crashed: {e}")
            failed += 1
    
    print("\n" + "=" * 60)
    print(f"Results: {passed} passed, {failed} failed")
    print("=" * 60)
    
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())