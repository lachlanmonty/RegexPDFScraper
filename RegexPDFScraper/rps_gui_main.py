from __future__ import with_statement
import os
import pandas as pd
import fitz
from rps_gui import Ui_RPS_GUI

from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc


class MainWindow(qtw.QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.file_names = None
        self.df_o = None
        self.ui = Ui_RPS_GUI()
        self.ui.setupUi(self)

        self.ui.open_file_button.clicked.connect(self.select_file)

        self.ui.regex_search_button.clicked.connect(self.regex_search)
        self.ui.save_button.clicked.connect(self.save_data_as_xlsx)

        # self.ui.save_button.clicked.connect(self.show_dialog)

    def regex_search(self):
        # search text in pdf
        def get_text(file, pattern):
            doc = fitz.open(file)
            text_comb = {}
            for pageNumber, page in enumerate(doc.pages(), start=1):
                text_comb[pageNumber] = page.get_text()
            df_txt = pd.DataFrame.from_dict(text_comb, orient="index", columns=["txt"])
            df_match = df_txt["txt"].str.extractall(pattern)
            df_match["file_path"] = file
            df_match["file_name"] = os.path.basename(os.path.normpath(file))
            return df_match

        # loop the get_text function
        def loop_get_text(filenames, pattern):
            df = []
            df = pd.concat([pd.DataFrame(get_text(i, pattern)) for i in filenames])
            df.reset_index(inplace=True)
            # df = df.drop_duplicates(subset=[0, "file_name"])
            columns = {"level_0": "page_no"}
            df = df.rename(columns=columns)
            return df

        try:
            self.df_o = loop_get_text(self.file_names, self.ui.regex_input.text())
            self.ui.regex_confirm.setText(f"{len(self.df_o)} match/s found")
        except:
            self.ui.regex_confirm.setText("No matches found")

    # def select_regex(self):
    #     # open file dialog
    #     self.regex = self.regex_input.toPlainText()
    #     self.ui.file_confirm.setText(f"{str(len(self.file_names))} file/s selected")

    def select_file(self):
        # open file dialog
        self.file_names, _ = qtw.QFileDialog.getOpenFileNames(self, "Open File", "", "pdf (*.pdf)")
        self.ui.file_confirm.setText(f"{str(len(self.file_names))} file/s selected")

    def save_data_as_xlsx(self):
        name = qtw.QFileDialog.getSaveFileName(self, "Save File", filter="*.xlsx")
        if name[0] == "":
            pass
        else:
            self.df_o.to_excel(name[0], index=False)

    # def show_dialog(self):
    #     dialog = qtw.QMessageBox(self)
    #     dialog.setText("File Saved Successfully")
    #     dialog.exec_()


if __name__ == "__main__":
    app = qtw.QApplication([])

    main_window = MainWindow()
    main_window.show()

    app.exec_()
