#!/usr/bin/env python3
"""
Shared utilities for Outlook tools.
"""

import base64


def encode_entry_id(hex_id):
    """
    Convert hex EntryID to base64 for token efficiency.
    
    Reduces token usage by ~30-40% (from ~35 tokens to ~20-25 tokens).
    140 hex chars → ~93 base64 chars.
    
    Args:
        hex_id: Hex string EntryID from Outlook
        
    Returns:
        Base64 encoded string, or original if encoding fails
    """
    if not hex_id:
        return None
    
    try:
        # Convert hex string to bytes, then to base64
        bytes_data = bytes.fromhex(hex_id)
        return base64.b64encode(bytes_data).decode('ascii')
    except Exception:
        # Return as-is if encoding fails
        return hex_id


def decode_entry_id(encoded_id):
    """
    Convert base64 EntryID back to hex, or pass through if already hex.
    
    Handles both formats for backward compatibility.
    
    Args:
        encoded_id: Either base64 or hex EntryID
        
    Returns:
        Hex string EntryID for Outlook API
    """
    if not encoded_id:
        return None
    
    # Check if it's already hex (140 chars of hex)
    if len(encoded_id) == 140:
        # Quick check for hex characters
        if all(c in '0123456789ABCDEFabcdef' for c in encoded_id):
            return encoded_id.upper()
    
    # Otherwise try to decode from base64
    try:
        decoded_bytes = base64.b64decode(encoded_id)
        return decoded_bytes.hex().upper()
    except Exception:
        # Return as-is if decoding fails (might be malformed)
        return encoded_id