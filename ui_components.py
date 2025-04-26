from PyQt5.QtWidgets import (QFrame, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QProgressBar, QTableWidget, QHeaderView, QWidget, QSizePolicy,
                             QToolBar, QAction, QToolButton, QMenu)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter, QPolygon, QColor
from PyQt5.QtCore import QPoint  # QPoint is often in QtCore instead of QtGui

from ui.custom_widgets import DraggableHeaderView


class UIComponents:
    """Class responsible for creating UI components."""

    def __init__(self, parent, fonts, current_columns):
        """Initialize with parent window and fonts."""
        self.parent = parent
        self.title_font = fonts['title']
        self.header_font = fonts['header']
        self.normal_font = fonts['normal']
        self.small_font = fonts['small']
        self.mono_font = fonts['mono']
        self.current_columns = current_columns

    def create_app_icon(self):
        """Create an application icon."""
        icon = QIcon()

        # Create a pixmap for the icon
        size = 128
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw the icon background (a rounded rectangle)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor('#5c6bc0'))
        painter.drawRoundedRect(0, 0, size, size, 15, 15)

        # Draw the document shape
        painter.setBrush(QColor('#ffffff'))
        doc_path = QPolygon([
            QPoint(size // 4, size // 6),
            QPoint(3 * size // 4, size // 6),
            QPoint(3 * size // 4, 5 * size // 6),
            QPoint(size // 4, 5 * size // 6)
        ])
        painter.drawPolygon(doc_path)

        # Draw some "text lines"
        painter.setPen(QColor('#5c6bc0'))
        line_y = size // 3
        line_spacing = size // 8
        for i in range(4):
            line_length = 2 * size // 5 if i % 2 == 0 else size // 3
            painter.drawLine(size // 3, line_y, size // 3 + line_length, line_y)
            line_y += line_spacing

        painter.end()

        icon.addPixmap(pixmap)
        return icon

    def create_toolbar(self):
        """Create a toolbar with actions and resources dropdown."""
        toolbar = QToolBar("Main Toolbar")
        toolbar.setIconSize(QSize(16, 16))
        toolbar.setMovable(False)
        self.parent.addToolBar(toolbar)

        # Open file action
        open_action = QAction("Open", self.parent)
        open_action.triggered.connect(self.parent.open_file)
        open_action.setStatusTip("Open an MXLIFF file")
        toolbar.addAction(open_action)

        # Export file action
        export_action = QAction("Export", self.parent)
        export_action.triggered.connect(self.parent.export_file)
        export_action.setStatusTip("Export the edited MXLIFF file")
        export_action.setEnabled(False)  # Disabled until a file is loaded
        toolbar.addAction(export_action)
        self.parent.export_action = export_action

        # Script Resources dropdown
        resources_menu = QMenu("Script Resources", self.parent)

        # Content Team Info action
        content_team_action = QAction("Content Team Info", self.parent)
        content_team_action.triggered.connect(self.parent.open_content_team_info)
        resources_menu.addAction(content_team_action)

        # Queries action
        queries_action = QAction("Queries", self.parent)
        queries_action.triggered.connect(self.parent.open_queries)
        resources_menu.addAction(queries_action)

        # Resources dropdown button
        resources_button = QToolButton(self.parent)
        resources_button.setText("Script Resources")
        resources_button.setMenu(resources_menu)
        resources_button.setPopupMode(QToolButton.InstantPopup)

        toolbar.addWidget(resources_button)

        # Add right-aligned spacer
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        toolbar.addWidget(spacer)

        # About action
        about_action = QAction("About", self.parent)
        about_action.triggered.connect(self.parent.show_about)
        toolbar.addAction(about_action)

        return toolbar

    def create_header_section(self):
        """Create the application header section."""
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_frame.setFixedHeight(120)

        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 10, 20, 10)
        header_layout.setSpacing(8)

        # Top row with title and version
        top_row = QHBoxLayout()

        # Title label
        title_label = QLabel('D4 Scripts Tool')
        title_label.setObjectName("titleLabel")
        title_label.setFont(self.title_font)
        top_row.addWidget(title_label)

        # Version label (right-aligned)
        version_label = QLabel("v1.5")  # Updated version to reflect the word count enhancement
        version_label.setObjectName("versionLabel")
        version_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        top_row.addWidget(version_label)

        header_layout.addLayout(top_row)

        # Middle section with file operations
        middle_row = QHBoxLayout()
        middle_row.setSpacing(10)

        # Buttons container
        buttons_container = QWidget()
        buttons_layout = QHBoxLayout(buttons_container)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(10)

        # Open file button
        open_button = QPushButton('Open MXLIFF File')
        open_button.setObjectName("primaryButton")
        open_button.clicked.connect(self.parent.open_file)
        open_button.setFixedSize(180, 40)
        buttons_layout.addWidget(open_button)
        self.parent.open_button = open_button

        # Export file button
        export_button = QPushButton('Export MXLIFF File')
        export_button.setObjectName("primaryButton")
        export_button.clicked.connect(self.parent.export_file)
        export_button.setFixedSize(180, 40)
        export_button.setEnabled(False)  # Disabled until a file is loaded
        buttons_layout.addWidget(export_button)
        self.parent.export_button = export_button

        # Add buttons container to middle row
        middle_row.addWidget(buttons_container)
        middle_row.addStretch(1)

        header_layout.addLayout(middle_row)

        # Bottom row with file info and progress
        bottom_row = QHBoxLayout()

        # File name label
        file_label = QLabel('No file selected')
        file_label.setObjectName("fileLabel")
        file_label.setWordWrap(True)
        bottom_row.addWidget(file_label)
        self.parent.file_label = file_label

        # Progress bar
        progress_bar = QProgressBar()
        progress_bar.setObjectName("progressBar")
        progress_bar.setVisible(False)
        progress_bar.setRange(0, 0)  # Indeterminate progress
        progress_bar.setFixedWidth(200)
        bottom_row.addWidget(progress_bar)
        self.parent.progress_bar = progress_bar

        header_layout.addLayout(bottom_row)

        return header_frame

    def create_table_panel(self):
        """Create the panel that contains the table."""
        content_frame = QFrame()
        content_frame.setObjectName("contentFrame")

        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(10, 10, 10, 10)
        content_layout.setSpacing(5)

        # Panel header with title
        panel_header = QHBoxLayout()

        panel_title = QLabel("Translation Data")
        panel_title.setObjectName("panelTitle")
        panel_title.setFont(self.header_font)
        panel_header.addWidget(panel_title)

        # Table stats
        table_stats = QLabel("0 entries")
        table_stats.setObjectName("tableStats")
        table_stats.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        panel_header.addWidget(table_stats)
        self.parent.table_stats = table_stats

        content_layout.addLayout(panel_header)

        # Create standard table
        table = QTableWidget()
        table.setObjectName("dataTable")
        table.setColumnCount(len(self.current_columns))
        table.setHorizontalHeaderLabels(self.current_columns)

        # Replace standard header with our draggable header
        header = DraggableHeaderView(Qt.Horizontal, table)
        table.setHorizontalHeader(header)
        self.parent.header = header

        # Configure table properties
        table.setShowGrid(True)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)
        table.setSelectionBehavior(QTableWidget.SelectRows)

        # Allow editing for the Target Text column
        table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)

        # Connect cell changed signal to handle edits
        table.cellChanged.connect(self.parent.on_cell_changed)

        # Connect selection change to display word count
        table.selectionModel().selectionChanged.connect(self.parent.on_selection_changed)

        # Enable text wrapping
        table.setWordWrap(True)

        # Allow rows to resize based on content
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # Hide vertical header (row numbers)
        table.verticalHeader().setVisible(False)

        content_layout.addWidget(table)
        self.parent.table = table

        return content_frame
