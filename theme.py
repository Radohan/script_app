class ThemeManager:
    """Manages application themes and styling."""

    @staticmethod
    def get_light_theme():
        """Returns the light theme color palette."""
        return {
            'app_bg': '#f9fafc',
            'header_bg': '#ffffff',
            'panel_bg': '#ffffff',
            'button_primary': '#5c6bc0',
            'button_hover': '#3f51b5',
            'button_secondary': '#eceff1',
            'button_secondary_hover': '#cfd8dc',
            'text_primary': '#263238',
            'text_secondary': '#607d8b',
            'border': '#e0e0e0',
            'table_header': '#f5f5f5',
            'table_alternate': '#f8f9fa',
            'table_selected': '#e8eaf6',
            'progress_bar': '#5c6bc0',
            'group_header': '#CAF1DE',
            'menu_label': '#FEF8DD',
            'female_key': '#FFE7C7',
            'diff_text': '#FF0000',  # Red text for differences
            'word_count_text': '#888888'  # Gray text for word count info
        }

    @staticmethod
    def get_dark_theme():
        """Returns the dark theme color palette."""
        return {
            'app_bg': '#263238',
            'header_bg': '#37474f',
            'panel_bg': '#37474f',
            'button_primary': '#7986cb',
            'button_hover': '#5c6bc0',
            'button_secondary': '#455a64',
            'button_secondary_hover': '#546e7a',
            'text_primary': '#eceff1',
            'text_secondary': '#b0bec5',
            'border': '#455a64',
            'table_header': '#455a64',
            'table_alternate': '#2b3f4b',
            'table_selected': '#3f51b5',
            'progress_bar': '#7986cb',
            'group_header': '#2e7d32',
            'menu_label': '#8d6e63',
            'female_key': '#ad1457',
            'diff_text': '#FF6B6B',  # Lighter red for dark mode
            'word_count_text': '#b0bec5'  # Light gray text for word count info in dark mode
        }

    @staticmethod
    def generate_stylesheet(theme):
        """Generates CSS stylesheet from theme dictionary."""
        return f"""
            QMainWindow, QDialog {{
                background-color: {theme['app_bg']};
                color: {theme['text_primary']};
            }}

            /* Header styling */
            #headerFrame {{
                background-color: {theme['header_bg']};
                border-bottom: 1px solid {theme['border']};
            }}

            #titleLabel {{
                color: {theme['text_primary']};
                font-size: 24px;
            }}

            #versionLabel {{
                color: {theme['text_secondary']};
            }}

            #fileLabel {{
                color: {theme['text_secondary']};
            }}

            /* Resources section styling - UPDATED: no border or background */
            #resourcesFrame {{
                background-color: transparent;
                border: none;
            }}

            #resourcesTitle {{
                color: #000000;  /* Black font */
                font-weight: bold;
            }}

            /* Button styling */
            #primaryButton {{
                background-color: {theme['button_primary']};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }}

            #primaryButton:hover {{
                background-color: {theme['button_hover']};
            }}

            #primaryButton:pressed {{
                background-color: {theme['button_hover']};
                padding: 9px 15px 7px 17px;
            }}

            #secondaryButton {{
                background-color: {theme['button_secondary']};
                color: {theme['text_primary']};
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
            }}

            #secondaryButton:hover {{
                background-color: {theme['button_secondary_hover']};
            }}

            /* Progress bar */
            #progressBar {{
                border: 1px solid {theme['border']};
                border-radius: 5px;
                text-align: center;
            }}

            #progressBar::chunk {{
                background-color: {theme['progress_bar']};
                width: 10px;
                margin: 0.5px;
            }}

            /* Table styling */
            #contentFrame {{
                background-color: {theme['panel_bg']};
                border: 1px solid {theme['border']};
                border-radius: 6px;
            }}

            #dataTable {{
                gridline-color: {theme['border']};
                background-color: {theme['panel_bg']};
                alternate-background-color: {theme['table_alternate']};
                font-size: 14px;
                border: none;
            }}

            #dataTable::item {{
                padding: 5px;
            }}

            QHeaderView::section {{
                background-color: {theme['table_header']};
                padding: 8px 5px;
                border: 1px solid {theme['border']};
                font-weight: bold;
                font-size: 12px;
                color: {theme['text_primary']};
            }}

            QHeaderView::section:hover {{
                background-color: {theme['button_secondary_hover']};
            }}

            #dataTable::item:selected {{
                background-color: {theme['table_selected']};
                color: {theme['text_primary']};
            }}

            /* Editing styles */
            QTableWidget QLineEdit {{
                background-color: {theme['panel_bg']};
                color: {theme['text_primary']};
                selection-background-color: {theme['button_primary']};
                border: 2px solid {theme['button_primary']};
                padding: 2px;
            }}

            /* Panel titles */
            #panelTitle {{
                color: {theme['text_primary']};
                font-size: 14px;
                font-weight: bold;
            }}

            #tableStats {{
                color: {theme['text_secondary']};
            }}

            /* Toolbar styling */
            QToolBar {{
                background-color: {theme['header_bg']};
                border-bottom: 1px solid {theme['border']};
            }}

            QToolBar QToolButton {{
                background-color: transparent;
                color: {theme['text_primary']};
                border: none;
                padding: 6px;
                margin: 2px;
            }}

            QToolBar QToolButton:hover {{
                background-color: {theme['button_secondary']};
                border-radius: 4px;
            }}

            /* Status bar */
            QStatusBar {{
                background-color: {theme['header_bg']};
                color: {theme['text_secondary']};
                border-top: 1px solid {theme['border']};
            }}

            /* Word count info style */
            .wordCountInfo {{
                color: {theme['word_count_text']};
                font-size: 8pt;
            }}
        """
