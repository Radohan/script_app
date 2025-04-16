from PyQt5.QtCore import QThread, pyqtSignal
import re
import traceback
import os
import codecs


class FileProcessingWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)

    def __init__(self, file_path, operation_type, parent=None):
        super().__init__()
        self.file_path = file_path
        self.operation_type = operation_type
        self.parent = parent
        self.data = {}
        self._cancelled = False

    def set_data(self, key, value):
        self.data[key] = value

    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            if self._cancelled:
                return

            if self.operation_type == 'parse_xml':
                self._parse_xml()
            elif self.operation_type == 'export_file':
                self._export_file()
            elif self.operation_type == 'process_document':
                self._process_document()
            else:
                self.error_signal.emit(f"Unknown operation type: {self.operation_type}")
        except Exception as e:
            import traceback
            trace = traceback.format_exc()
            self.error_signal.emit(f"Error in worker thread: {str(e)}\n{trace}")

    def _parse_xml(self):
        """Worker thread method to parse XML."""
        try:
            # Report progress
            self.progress_signal.emit(10, "Reading file...")

            # Read file in chunks to avoid memory issues
            with open(self.file_path, 'rb') as f:
                # Get file size
                f.seek(0, 2)
                file_size = f.tell()
                f.seek(0)

                # Read in chunks of 1MB
                chunk_size = 1024 * 1024
                content = b''

                bytes_read = 0
                while True:
                    if self._cancelled:
                        return

                    chunk = f.read(chunk_size)
                    if not chunk:
                        break

                    content += chunk
                    bytes_read += len(chunk)

                    # Report progress
                    progress = int((bytes_read / file_size) * 50)  # 50% for reading
                    self.progress_signal.emit(progress,
                                              f"Reading file... ({bytes_read / 1024 / 1024:.1f}MB / {file_size / 1024 / 1024:.1f}MB)")

            # Decode content
            self.progress_signal.emit(50, "Decoding content...")
            xml_content = content.decode('utf-8', errors='ignore')

            # Parse XML
            self.progress_signal.emit(60, "Parsing XML...")
            from utils.xml_parser import XMLParser
            processed_data = XMLParser.parse_xml(xml_content, print)

            # Emit results
            self.progress_signal.emit(100, "Processing complete")
            self.finished_signal.emit({
                'processed_data': processed_data,
                'original_xml_content': xml_content,
                'file_path': self.file_path
            })

        except Exception as e:
            import traceback
            trace = traceback.format_exc()
            self.error_signal.emit(f"Error parsing XML: {str(e)}\n{trace}")

    def _export_file(self):
        """Worker thread method to export file."""
        try:
            self.progress_signal.emit(10, "Preparing export...")

            # Get data from parent
            processed_data = self.data.get('processed_data', [])
            original_xml_content = self.data.get('original_xml_content', '')

            # Update XML content
            self.progress_signal.emit(30, "Updating XML content...")
            from utils.xml_parser import XMLParser
            updated_xml = XMLParser.update_xml_content(
                original_xml_content,
                processed_data,
                print
            )

            # Write to file
            self.progress_signal.emit(70, "Writing to file...")
            import codecs
            with codecs.open(self.file_path, 'w', 'utf-8') as f:
                f.write(updated_xml)

            # Emit results
            self.progress_signal.emit(100, "Export complete")
            self.finished_signal.emit({
                'success': True,
                'file_path': self.file_path
            })

        except Exception as e:
            import traceback
            trace = traceback.format_exc()
            self.error_signal.emit(f"Error exporting file: {str(e)}\n{trace}")

    def _process_document(self):
        """Worker thread method to process document."""
        try:
            self.progress_signal.emit(10, "Initializing document parser...")

            # Get document parser from parent if available
            document_parser = getattr(self.parent, 'document_parser', None)
            if not document_parser:
                from utils.document_parser import DocumentParser
                document_parser = DocumentParser(self.parent)

            # Parse document
            self.progress_signal.emit(20, "Parsing document...")
            success = document_parser.parse_document(self.file_path)
            if not success:
                self.error_signal.emit("Failed to parse document.")
                return

            # Get conversation tables
            self.progress_signal.emit(50, "Finding conversation tables...")
            conversation_tables = document_parser.get_conversation_tables()
            if not conversation_tables:
                self.error_signal.emit("No conversation tables found in document.")
                return

            # Match content with MXLIFF data
            self.progress_signal.emit(70, "Matching content with MXLIFF data...")
            processed_data = self.data.get('processed_data', [])
            match_results = document_parser.match_content_with_mxliff(processed_data)

            # Emit results
            self.progress_signal.emit(100, "Document processing complete")
            self.finished_signal.emit({
                'success': True,
                'tables': conversation_tables,
                'match_results': match_results
            })

        except Exception as e:
            import traceback
            trace = traceback.format_exc()
            self.error_signal.emit(f"Error processing document: {str(e)}\n{trace}")