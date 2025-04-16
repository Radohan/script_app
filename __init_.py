# This file allows the ui directory to be imported as a package

from ui.main_window import MXLIFFParser
from ui.custom_widgets import DraggableHeaderView, DiffHighlighter, TranslationDiffDialog
from ui.ui_components import UIComponents
from ui.theme import ThemeManager

# Export the main classes
__all__ = [
    'MXLIFFParser',
    'DraggableHeaderView',
    'DiffHighlighter',
    'TranslationDiffDialog',
    'UIComponents',
    'ThemeManager'
]
