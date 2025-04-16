from PyQt5.QtWidgets import QHeaderView, QDialog, QTextEdit, QLabel, QVBoxLayout, QHBoxLayout
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor, QTextCharFormat, QSyntaxHighlighter

class DraggableHeaderView(QHeaderView):
    """Custom header view that prevents column reordering."""

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsMovable(False)
        self.setDragEnabled(False)
        self.setSectionsClickable(True)


class DiffHighlighter(QSyntaxHighlighter):
    """A syntax highlighter that highlights differences in text."""

    def __init__(self, document, diff_words):
        super().__init__(document)
        self.diff_words = diff_words
        self.diff_format = QTextCharFormat()
        self.diff_format.setForeground(QColor("red"))
        self.diff_format.setFontWeight(QFont.Bold)

    def highlightBlock(self, text):
        """Highlight words that are in the diff_words list."""
        for word in self.diff_words:
            if not word:
                continue

            # Find all occurrences of the word
            start_index = 0
            while start_index < len(text):
                index = text.find(word, start_index)
                if index == -1:
                    break

                # Make sure it's a whole word
                is_whole_word = True
                if index > 0 and text[index - 1].isalnum():
                    is_whole_word = False
                end_index = index + len(word)
                if end_index < len(text) and text[end_index].isalnum():
                    is_whole_word = False

                if is_whole_word:
                    self.setFormat(index, len(word), self.diff_format)

                start_index = index + len(word)


class TranslationDiffDialog(QDialog):
    """Dialog for viewing translation differences side by side."""

    def __init__(self, parent, male_text, female_text, diff_words):
        super().__init__(parent)
        self.setWindowTitle("Translation Comparison")
        self.resize(800, 400)

        # Create layout
        layout = QVBoxLayout(self)

        # Labels
        header_layout = QHBoxLayout()
        male_label = QLabel("Standard Version:")
        male_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        female_label = QLabel("Female Version:")
        female_label.setFont(QFont("Segoe UI", 10, QFont.Bold))

        header_layout.addWidget(male_label)
        header_layout.addWidget(female_label)
        layout.addLayout(header_layout)

        # Text areas
        texts_layout = QHBoxLayout()

        # Male text
        self.male_text_edit = QTextEdit()
        self.male_text_edit.setReadOnly(True)
        self.male_text_edit.setPlainText(male_text)

        # Female text with highlighting
        self.female_text_edit = QTextEdit()
        self.female_text_edit.setReadOnly(True)
        self.female_text_edit.setPlainText(female_text)

        # Apply highlighting to female text
        self.highlighter = DiffHighlighter(self.female_text_edit.document(), diff_words)

        texts_layout.addWidget(self.male_text_edit)
        texts_layout.addWidget(self.female_text_edit)
        layout.addLayout(texts_layout)