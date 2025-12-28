"""
Utility functions for text cleaning and validation.
"""

import re
from typing import Optional
from pathlib import Path


def is_section_heading(text: str) -> bool:
    """
    Detect all-caps section headings (e.g., 'COBALT', 'LITHIUM').
    
    Section headings are identified by:
    - All uppercase letters
    - May contain spaces, parentheses, slashes, hyphens, numbers
    - Minimum length of 3 characters
    
    Args:
        text: String to check
        
    Returns:
        True if text matches section heading pattern, False otherwise
        
    Examples:
        >>> is_section_heading("COBALT")
        True
        >>> is_section_heading("RENEWABLE ENERGY")
        True
        >>> is_section_heading("Introduction")
        False
        >>> is_section_heading("CO")
        False
    """
    if not text:
        return False
    
    pattern = re.compile(r"^[A-Z][A-Z\s\(\)\/\-0-9]*$")
    return bool(pattern.match(text.strip())) and len(text.strip()) > 2

def clean_numeric_like(s: Optional[str]) -> str:
    """
    Minimal cleaning for extraction phase.
    
    Only removes formatting that breaks Excel imports:
    - Multi-line content (keeps first line)
    - Excess whitespace
    
    Preserves:

    - Superscripts (handled in downstream cleaning)
    - Country name variations
    - All original text
    
    
    """
    if s is None or not isinstance(s, str):
        return ""
    
    s = s.strip()
    
    # Handle obvious missing values
    if s.upper() in {"NA", "N/A", "XX", "W"}:
        return ""
    
    # ONLY fix: Multi-line cells (Excel doesn't like these)
    if '\n' in s or '\r' in s:
        lines = [line.strip() for line in s.split('\n') if line.strip()]
        s = lines[0] if lines else ""
    
    # Remove tabs (Excel formatting issue)
    s = s.replace('\t', ' ')
    
    # That's it! Keep everything else as-is
    return s.strip()

def clean_sheet_name(name: str, max_length: int = 31) -> str:
    """
    Create valid Excel sheet names.
    
    Excel sheet name requirements:
    - Maximum 31 characters
    - Cannot contain: [ ] : * ? / \\
    - Cannot be empty
    
    Args:
        name: Proposed sheet name
        max_length: Maximum length (default: 31 for Excel)
        
    Returns:
        Valid Excel sheet name
        
    Examples:
        >>> clean_sheet_name("COBALT/LITHIUM")
        'COBALTLITHIUM'
        >>> clean_sheet_name("A" * 40)
        'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    """
    # Remove invalid characters
    name = re.sub(r"[\[\]\:\*\?\/\\]", "", name).strip()
    
    # Truncate to max length
    name = name[:max_length]
    
    # Ensure not empty
    return name if name else "Sheet"


def validate_file_path(path, must_exist: bool = True) -> bool:
    """
    Validate file path.
    
    Args:
        path: Path object or string
        must_exist: If True, check that file exists
        
    Returns:
        True if valid, False otherwise
    """
    try:
        path = Path(path)
        
        if must_exist:
            return path.exists() and path.is_file()
        else:
            return True
            
    except (TypeError, ValueError):
        return False