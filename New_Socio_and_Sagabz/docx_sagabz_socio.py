import os
from docx_helper import Docx_helper
from column_constants import MASHOV_SAGZAB, SOCIOMETRY

class Docx_sagzab(Docx_helper):
    def __init__(self, file_format_path, word_output_dir):
        super().__init__(file_format_path, word_output_dir)
        self.new_columns = MASHOV_SAGZAB

    def is_values(self, category: str) -> bool:
        return False

    def get_columns_names(self):
        column_names = self.new_columns
        return column_names


class Docx_Socio(Docx_helper):
    def __init__(self, file_format_path, word_output_dir):
        super().__init__(file_format_path, word_output_dir)
        self.new_columns = SOCIOMETRY

    def is_values(self, category: str) -> bool:
        if category in self.new_columns[-8:-2]:
            return True
        return False

    def get_columns_names(self):
        column_names = self.new_columns
        return column_names