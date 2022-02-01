from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtPrintSupport import *
from PyQt5.QtCore import *
import sys
import docx2txt

class BSWord(QMainWindow):
    def __init__(self):
        super(BSWord, self).__init__()
        self.editor = QTextEdit()
        self.setCentralWidget(self.editor)
        self.font_size_box = QSpinBox()
        self.showMaximized()
        self.setWindowTitle('SV WORD')
        self.create_tool_bar()
        self.create_menu_bar()
        font = QFont('Times, 12')
        self.editor.setFont(font)
        self.editor.setFontPointSize(20)

        self.path = ''

    def create_menu_bar(self):
        menu_bar = QMenuBar()

        app_icon = menu_bar.addMenu(QIcon("doc_icon.png"), "icon")

        file_menu = QMenu('File', self)
        menu_bar.addMenu(file_menu)

        open_action = QAction('Open', self)
        open_action.triggered.connect(self.file_open)
        file_menu.addAction(open_action)

        rename_action = QAction('Rename', self)
        rename_action.triggered.connect(self.file_saveas)
        file_menu.addAction(rename_action)

        save_action = QAction('Save', self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        save_as_pdf_action = QAction('Save as PDF', self)
        save_as_pdf_action.triggered.connect(self.save_as_pdf)
        file_menu.addAction(save_as_pdf_action)

        edit_menu = QMenu('Edit', self)
        menu_bar.addMenu(edit_menu)

        paste_action = QAction('Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)

        clear_action = QAction('Clear', self)
        clear_action.triggered.connect(self.editor.clear)
        edit_menu.addAction(clear_action)

        select_action = QAction('Select All', self)
        select_action.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(select_action)

        view_menu = QMenu('View', self)
        menu_bar.addMenu(view_menu)

        fullscr_action = QAction('Full Screen View', self)
        fullscr_action.triggered.connect(lambda : self.showFullScreen())
        view_menu.addAction(fullscr_action)

        normscr_action = QAction('Normal View', self)
        normscr_action.triggered.connect(lambda : self.showNormal())
        view_menu.addAction(normscr_action)

        minscr_action = QAction('Minimize', self)
        minscr_action.triggered.connect(lambda : self.showMinimized())
        view_menu.addAction(minscr_action)

        self.setMenuBar(menu_bar)


    def create_tool_bar(self):
        tool_bar = QToolBar()

        undo_action = QAction(QIcon('undo.png'), 'Undo', self)
        undo_action.triggered.connect(self.editor.undo)
        tool_bar.addAction(undo_action)

        redo_action = QAction(QIcon('redo.png'), 'Redo', self)
        redo_action.triggered.connect(self.editor.redo)
        tool_bar.addAction(redo_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        cut_action = QAction(QIcon('cut.png'), 'Cut', self)
        cut_action.triggered.connect(self.editor.cut)
        tool_bar.addAction(cut_action)

        copy_action = QAction(QIcon('copy.png'), 'Copy', self)
        copy_action.triggered.connect(self.editor.copy)
        tool_bar.addAction(copy_action)

        paste_action = QAction(QIcon('paste.png'), 'Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        tool_bar.addAction(paste_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        self.font_combo = QComboBox(self)
        self.font_combo.addItems(["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"])
        self.font_combo.activated.connect(self.set_font)      # connect with function
        tool_bar.addWidget(self.font_combo) 

        tool_bar.addSeparator()
    
        self.font_size_box.setValue(20)
        self.font_size_box.valueChanged.connect(self.set_font_size)
        tool_bar.addWidget(self.font_size_box)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        bold_action = QAction(QIcon('bold.png'), 'Bold', self)
        bold_action.triggered.connect(self.bold_text)
        tool_bar.addAction(bold_action)

        underline_action = QAction(QIcon('underline.png'), 'Underline', self)
        underline_action.triggered.connect(self.underline_text)
        tool_bar.addAction(underline_action)

        italic_action = QAction(QIcon('italic.png'), 'Italic', self)
        italic_action.triggered.connect(self.italic_text)
        tool_bar.addAction(italic_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        left_alignment_action = QAction(QIcon("left-align.png"), 'Align Left', self)
        left_alignment_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignLeft))
        tool_bar.addAction(left_alignment_action)

        right_alignment_action = QAction(QIcon("right-align.png"), 'Align Right', self)
        right_alignment_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignRight))
        tool_bar.addAction(right_alignment_action)

        justification_action = QAction(QIcon("justification.png"), 'Center/Justify', self)
        justification_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignCenter))
        tool_bar.addAction(justification_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        zoom_in_action = QAction(QIcon("zoom-in.png"), 'Zoom in', self)
        zoom_in_action.triggered.connect(self.editor.zoomIn)
        tool_bar.addAction(zoom_in_action)

        zoom_out_action = QAction(QIcon("zoom-out.png"), 'Zoom out', self)
        zoom_out_action.triggered.connect(self.editor.zoomOut)
        tool_bar.addAction(zoom_out_action)

        self.addToolBar(tool_bar)

    def italic_text(self):
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not(state))

    def underline_text(self):
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))
    
    def bold_text(self):
        if self.editor.fontWeight() != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)


    def set_font(self):
        font = self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(font))

    def set_font_size(self):
        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)

    def save_as_pdf(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Export PDF', None, 'PDF Files (*.pdf)')
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        self.editor.document().print_(printer)

    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Text Documents(*.text); Text Documents(*.txt); All Files(*.*)")

        try:
            text = docx2txt.process(self.path)

        except Exception as e:
            print(e)
        
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        print(self.path)
        if self.path == '':
            self.file_saveas()

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)
    
    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "Text documents (*.text); Text documents (*.txt); All files (*.*)")

        if self.path == '':
            return  

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)


app = QApplication(sys.argv)
window = BSWord()
window.show()
sys.exit(app.exec_())