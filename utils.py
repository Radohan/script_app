import re
import difflib

def extract_main_key(full_key):
    """Extract the main part of the key (before slash or last period)."""
    if not full_key:
        return ''
    slash_index = full_key.find('/')
    if slash_index != -1:
        return full_key[:slash_index]
    last_dot_index = full_key.rfind('.')
    if last_dot_index != -1:
        return full_key[:last_dot_index]
    return full_key

def extract_line_number(item):
    """
    Extract the line number from the item's key or note text.
    Looks for patterns like 'Line_1', 'Line:2', etc.
    """
    key = item.get('key', '')
    note_text = item.get('note_text', '')

    # Try extracting from key first
    line_match = re.search(r'Line[_:](\d+)', key, re.IGNORECASE)
    if line_match:
        return int(line_match.group(1))

    # If not in key, try note text
    line_match = re.search(r'Line[_:](\d+)', note_text, re.IGNORECASE)
    if line_match:
        return int(line_match.group(1))

    # Fallback to order value if no explicit line number found
    return item.get('order_value', 9999)

def extract_order_value(note_text):
    """Extract the Order value from the key note."""
    if not note_text:
        return 9999  # Default high value for items without order
    order_match = re.search(r'Order:\s*(\d+)', note_text, re.IGNORECASE)
    return int(order_match.group(1)) if order_match else 9999


def has_comments(item):
    """
    Check if an item has any type of comments.
    Now checks for Developer Comment, regular Comment, and CoT Comment.
    """
    note_text = item.get('note_text', '')
    if not note_text:
        return False

    # Check for any type of comment
    return bool(re.search(r'(Developer Comment:|Comment:|CoT Comment:)', note_text))


def get_comment_text(item):
    """
    Extract comment text from an item, now supporting multiple comment types.
    Returns a combined string of all comment types.
    """
    note_text = item.get('note_text', '')
    if not note_text:
        return ""

    comments = []

    # Extract Developer Comment if present
    dev_match = re.search(r'Developer Comment:(.+?)(?=\n|$)', note_text, re.DOTALL)
    if dev_match:
        comments.append(f"Developer Comment: {dev_match.group(1).strip()}")

    # Extract regular Comment if present
    comment_match = re.search(r'Comment:(?!.*CoT)(.+?)(?=\n|$)', note_text, re.DOTALL)
    if comment_match:
        comments.append(f"Comment: {comment_match.group(1).strip()}")

    # Extract CoT Comment if present
    cot_match = re.search(r'CoT Comment:(.+?)(?=\n|$)', note_text, re.DOTALL)
    if cot_match:
        comments.append(f"CoT Comment: {cot_match.group(1).strip()}")

    # Join all comment types with line breaks
    return "\n".join(comments)

def natural_sort_key(k):
    """
    Custom sorting key that prioritizes Main quest dialogues and
    sorts numerically based on numbers in the key.
    """
    import re

    # Check if it's a Main quest
    is_main_quest = 'Main' in k

    # Extract numbers
    numbers = re.findall(r'\d+', k)

    # Convert found numbers to integers
    numeric_parts = [int(num) for num in numbers] if numbers else []

    # Return a tuple that prioritizes Main quests and then sorts numerically
    return (not is_main_quest, numeric_parts)

def find_text_differences(text1, text2):
    """
    Find word-level differences between two texts using difflib.
    Returns a list of words that differ in text2 compared to text1.
    """
    if text1 == text2:
        return []

    # Handle None values
    text1 = text1 or ""
    text2 = text2 or ""

    try:
        # Use difflib to find differences
        differ = difflib.Differ()
        diff = list(differ.compare(text1.split(), text2.split()))

        # Extract words that are in text2 but not in text1
        diff_words = []
        for word in diff:
            if word.startswith('+ '):  # Added in text2
                diff_words.append(word[2:])

        return diff_words
    except Exception as e:
        # Log the error and return empty list
        print(f"Error in find_text_differences: {str(e)}")
        return []