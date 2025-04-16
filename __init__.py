# This file allows the utils directory to be imported as a package

from utils.utils import (
    extract_main_key,
    extract_line_number,
    extract_order_value,
    has_comments,
    get_comment_text,
    natural_sort_key,
    find_text_differences
)

from utils.xml_parser import XMLParser
from utils.document_parser import DocumentParser

# Export utility functions and classes
__all__ = [
    'extract_main_key',
    'extract_line_number',
    'extract_order_value',
    'has_comments',
    'get_comment_text',
    'natural_sort_key',
    'find_text_differences',
    'XMLParser',
    'DocumentParser'
]