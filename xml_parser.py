import re
import traceback
from utils.utils import extract_order_value


class XMLParser:
    """Handles parsing and exporting MXLIFF XML files."""

    @staticmethod
    def parse_xml(xml_content, logger=None):
        """Parse XML directly using regex approach for MXLIFF files."""
        if logger:
            logger("Parsing XML using direct regex approach...")

        # First, find all group elements to properly associate context with trans-units
        group_pattern = r'<group[^>]*>.*?</group>'
        groups = re.findall(group_pattern, xml_content, re.DOTALL)

        if logger:
            logger(f"Found {len(groups)} group elements in the file")

        # Process all groups and their trans-units
        processed_data = []

        for group_idx, group in enumerate(groups):
            try:
                # Extract group ID
                group_id_match = re.search(r'<group\s+id="([^"]*)"', group)
                group_id = group_id_match.group(1) if group_id_match else f"group_{group_idx}"

                # Extract context-group for this group (if any)
                context_group_pattern = r'<context-group[^>]*>(.*?)</context-group>'
                context_group_match = re.search(context_group_pattern, group, re.DOTALL)

                # Default context information for this group
                group_key = ""
                group_note_text = ""

                if context_group_match:
                    context_group_content = context_group_match.group(1)

                    # Extract key
                    key_pattern = r'<context\s+context-type="x-key"[^>]*>(.*?)</context>'
                    key_match = re.search(key_pattern, context_group_content, re.DOTALL)
                    if key_match:
                        group_key = key_match.group(1).strip()

                    # Extract notes
                    note_pattern = r'<context\s+context-type="x-key-note"[^>]*>(.*?)</context>'
                    note_match = re.search(note_pattern, context_group_content, re.DOTALL)
                    if note_match:
                        group_note_text = note_match.group(1).strip()

                # Find all trans-units within this group
                trans_unit_pattern = r'<trans-unit[^>]*>.*?</trans-unit>'
                trans_units = re.findall(trans_unit_pattern, group, re.DOTALL)

                if logger:
                    logger(f"Group {group_id} contains {len(trans_units)} trans-units")

                # Process trans-units in this group
                for trans_idx, trans_unit in enumerate(trans_units):
                    # Extract source and target
                    source_pattern = r'<source[^>]*>(.*?)</source>'
                    target_pattern = r'<target[^>]*>(.*?)</target>'

                    source_match = re.search(source_pattern, trans_unit, re.DOTALL)
                    target_match = re.search(target_pattern, trans_unit, re.DOTALL)

                    source_text = source_match.group(1).strip() if source_match else ""
                    target_text = target_match.group(1).strip() if target_match else ""

                    # Extract trans-unit ID for better tracking
                    trans_id_match = re.search(r'<trans-unit\s+id="([^"]*)"', trans_unit)
                    trans_id = trans_id_match.group(1) if trans_id_match else f"trans_{trans_idx}"

                    # Check if this trans-unit has its own context information
                    unit_context_group_match = re.search(context_group_pattern, trans_unit, re.DOTALL)

                    # Variables to store context information for this specific trans-unit
                    unit_key = group_key
                    unit_note_text = group_note_text

                    if unit_context_group_match:
                        # This trans-unit has its own context group, override the group-level context
                        unit_context_content = unit_context_group_match.group(1)

                        # Extract key
                        unit_key_match = re.search(key_pattern, unit_context_content, re.DOTALL)
                        if unit_key_match:
                            unit_key = unit_key_match.group(1).strip()

                        # Extract notes
                        unit_note_match = re.search(note_pattern, unit_context_content, re.DOTALL)
                        if unit_note_match:
                            unit_note_text = unit_note_match.group(1).strip()

                    # Extract metadata from note text
                    speaker = ""
                    speaker_target = ""
                    speaker_gender = ""
                    player_class = ""
                    player_gender = ""
                    order_value = 9999

                    if unit_note_text:
                        # Extract speaker information
                        speaker_match = re.search(r'Speaker:\s*([^\n]+)', unit_note_text)
                        speaker_target_match = re.search(r'Target:\s*([^\n]+)', unit_note_text)
                        speaker_gender_match = re.search(r'Speaker Gender:\s*([^\n]+)', unit_note_text)
                        player_class_match = re.search(r'Class:\s*([^\n]+)', unit_note_text)
                        player_gender_match = re.search(r'Player Gender:\s*([^\n]+)', unit_note_text)

                        if speaker_match:
                            speaker = speaker_match.group(1).strip()
                        if speaker_target_match:
                            speaker_target = speaker_target_match.group(1).strip()
                        if speaker_gender_match:
                            speaker_gender = speaker_gender_match.group(1).strip()
                        if player_class_match:
                            player_class = player_class_match.group(1).strip()
                        if player_gender_match:
                            player_gender = player_gender_match.group(1).strip()

                        # Additional specific patterns to ensure we capture the gender info
                        if not speaker_gender:
                            gender_match = re.search(r'Gender:\s*([^,\n]+)', unit_note_text)
                            if gender_match:
                                speaker_gender = gender_match.group(1).strip()

                        # Look for "speaking to:" pattern which might indicate player gender
                        speaking_to_match = re.search(r'speaking to:\s*([^,\n]+)', unit_note_text)
                        if speaking_to_match and not speaker_target:
                            speaker_target = speaking_to_match.group(1).strip()

                        # Extract order value
                        order_value = extract_order_value(unit_note_text)

                    # Create a record
                    processed_data.append({
                        'index': len(processed_data),
                        'group_id': group_id,
                        'trans_id': trans_id,
                        'source_text': source_text,
                        'target_text': target_text,
                        'original_target_text': target_text,  # Store original for change detection
                        'key': unit_key,
                        'speaker': speaker,
                        'speaker_target': speaker_target,
                        'speaker_gender': speaker_gender,
                        'player_class': player_class,
                        'player_gender': player_gender,
                        'order_value': order_value,
                        'note_text': unit_note_text,
                        'is_menulabel': 'MenuLabel' in unit_key  # Flag for MenuLabel entries
                    })

            except Exception as e:
                if logger:
                    logger(f"Error processing group {group_idx}: {str(e)}")
                    logger(traceback.format_exc())

        if logger:
            logger(f"Total processed items: {len(processed_data)}")

        return processed_data

    @staticmethod
    def update_xml_content(original_xml_content, processed_data, logger=None):
        """Update the original XML content with the edited translations."""
        if not original_xml_content:
            raise ValueError("No original XML content available")

        if not processed_data:
            raise ValueError("No processed data available")

        # Dictionary to collect all translations by key
        key_to_translation = {}

        # Dictionary to track which translations were edited
        edited_translations = {}

        # Extract all translations from processed data
        for data in processed_data:
            if not data['is_header'] and 'item' in data:
                item = data['item']
                key = item.get('key', '')
                if key:
                    target_text = item.get('target_text', '')
                    original_text = item.get('original_target_text', '')
                    key_to_translation[key] = target_text

                    # Track which ones were actually changed
                    if target_text != original_text:
                        edited_translations[key] = {
                            'new': target_text,
                            'original': original_text
                        }

        if logger:
            logger(f"Found {len(edited_translations)} edited translations out of {len(key_to_translation)} total")

        # If nothing was edited, return the original content
        if not edited_translations:
            if logger:
                logger("No translations were edited, returning original XML")
            return original_xml_content

        # Create a working copy of the XML content
        updated_xml = original_xml_content

        # We'll use two approaches to ensure all edits are captured
        # 1. Directly find and replace each trans-unit
        # 2. Fall back to searching within group elements if needed

        # First approach: Find all trans-units and update directly
        trans_unit_pattern = r'(<trans-unit[^>]*>.*?</trans-unit>)'
        trans_units = re.findall(trans_unit_pattern, updated_xml, re.DOTALL)

        if logger:
            logger(f"Found {len(trans_units)} total trans-units in XML")

        # Dictionary to track which keys we've successfully updated
        updated_keys = set()

        # Process each trans-unit
        for trans_unit in trans_units:
            # Try to extract key from context-group
            context_group_pattern = r'<context-group[^>]*>(.*?)</context-group>'
            context_group_match = re.search(context_group_pattern, trans_unit, re.DOTALL)

            key = None
            if context_group_match:
                context_content = context_group_match.group(1)
                key_pattern = r'<context\s+context-type="x-key"[^>]*>(.*?)</context>'
                key_match = re.search(key_pattern, context_content, re.DOTALL)
                if key_match:
                    key = key_match.group(1).strip()

            # If we found a key and it's one we edited
            if key and key in edited_translations:
                # Extract the target tag
                target_pattern = r'(<target[^>]*>)(.*?)(</target>)'
                target_match = re.search(target_pattern, trans_unit, re.DOTALL)

                if target_match:
                    # Create updated trans-unit with new translation
                    new_text = edited_translations[key]['new']
                    updated_trans_unit = trans_unit.replace(
                        target_match.group(0),
                        f"{target_match.group(1)}{new_text}{target_match.group(3)}"
                    )

                    # Use a very specific replacement to avoid accidental matches
                    # We wrap the trans-unit with unique markers and then replace the whole block
                    marker_start = f"<!-- XMLUPDATE_START_{id(trans_unit)} -->"
                    marker_end = f"<!-- XMLUPDATE_END_{id(trans_unit)} -->"

                    marked_original = f"{marker_start}{trans_unit}{marker_end}"
                    marked_updated = f"{marker_start}{updated_trans_unit}{marker_end}"

                    # First add the markers
                    updated_xml = updated_xml.replace(trans_unit, marked_original)
                    # Then replace the marked block
                    updated_xml = updated_xml.replace(marked_original, marked_updated)

                    # Track that we've updated this key
                    updated_keys.add(key)

        # Check if all edits were applied
        missing_keys = set(edited_translations.keys()) - updated_keys

        if missing_keys:
            if logger:
                logger(f"Warning: {len(missing_keys)} edited translations could not be directly applied")
                logger(f"Attempting second approach for keys: {', '.join(list(missing_keys)[:5])}" +
                       (f"... and {len(missing_keys) - 5} more" if len(missing_keys) > 5 else ""))

            # Second approach: Find keys within groups
            # This is more complex but catches cases where context-group structure is different
            group_pattern = r'(<group[^>]*>.*?</group>)'
            groups = re.findall(group_pattern, updated_xml, re.DOTALL)

            for group in groups:
                # Extract trans-units from this group
                group_trans_units = re.findall(trans_unit_pattern, group, re.DOTALL)
                group_modified = False
                updated_group = group

                for trans_unit in group_trans_units:
                    # Try all possible patterns to extract key
                    key = None

                    # Pattern 1: Direct context-group in trans-unit
                    if not key:
                        context_group_match = re.search(context_group_pattern, trans_unit, re.DOTALL)
                        if context_group_match:
                            context_content = context_group_match.group(1)
                            key_match = re.search(r'<context\s+context-type="x-key"[^>]*>(.*?)</context>',
                                                  context_content, re.DOTALL)
                            if key_match:
                                key = key_match.group(1).strip()

                    # Pattern 2: Key in trans-unit attributes
                    if not key:
                        key_attr_match = re.search(r'<trans-unit[^>]*key="([^"]*)"', trans_unit, re.DOTALL)
                        if key_attr_match:
                            key = key_attr_match.group(1).strip()

                    # Pattern 3: Key in group context and trans-unit id match
                    if not key:
                        # Get context from group
                        group_context_match = re.search(context_group_pattern, group, re.DOTALL)
                        if group_context_match:
                            group_context = group_context_match.group(1)
                            group_key_match = re.search(r'<context\s+context-type="x-key"[^>]*>(.*?)</context>',
                                                        group_context, re.DOTALL)
                            if group_key_match:
                                base_key = group_key_match.group(1).strip()

                                # Get trans-unit id
                                trans_id_match = re.search(r'<trans-unit[^>]*id="([^"]*)"', trans_unit,
                                                           re.DOTALL)
                                if trans_id_match:
                                    trans_id = trans_id_match.group(1).strip()

                                    # Check if trans_id appears in any of our missing keys
                                    for missing_key in missing_keys:
                                        if trans_id in missing_key or missing_key in trans_id:
                                            key = missing_key
                                            break

                    # If we found a key and it's one we're looking for
                    if key and key in missing_keys:
                        # Extract the target tag
                        target_pattern = r'(<target[^>]*>)(.*?)(</target>)'
                        target_match = re.search(target_pattern, trans_unit, re.DOTALL)

                        if target_match:
                            # Create updated trans-unit with new translation
                            new_text = edited_translations[key]['new']
                            updated_trans_unit = trans_unit.replace(
                                target_match.group(0),
                                f"{target_match.group(1)}{new_text}{target_match.group(3)}"
                            )

                            # Update in the group
                            updated_group = updated_group.replace(trans_unit, updated_trans_unit)
                            group_modified = True

                            # Track that we've updated this key
                            updated_keys.add(key)

                # If this group was modified, update it in the XML
                if group_modified:
                    # Again use specific markers to avoid accidental replacements
                    marker_start = f"<!-- XMLUPDATE_GROUP_START_{id(group)} -->"
                    marker_end = f"<!-- XMLUPDATE_GROUP_END_{id(group)} -->"

                    marked_original = f"{marker_start}{group}{marker_end}"
                    marked_updated = f"{marker_start}{updated_group}{marker_end}"

                    # First add the markers
                    updated_xml = updated_xml.replace(group, marked_original)
                    # Then replace the marked block
                    updated_xml = updated_xml.replace(marked_original, marked_updated)

        # Final check for any remaining missing keys
        final_missing = set(edited_translations.keys()) - updated_keys
        if final_missing and logger:
            logger(f"Warning: {len(final_missing)} edited translations could not be applied to the XML")
            logger(f"Missing keys: {', '.join(list(final_missing)[:5])}" +
                   (f"... and {len(final_missing) - 5} more" if len(final_missing) > 5 else ""))

        if logger:
            logger(f"Successfully updated {len(updated_keys)} out of {len(edited_translations)} edited translations")

        # Clean up any markers that might have been left in the XML
        cleaned_xml = re.sub(r'<!-- XMLUPDATE_[^>]*-->', '', updated_xml)

        return cleaned_xml
