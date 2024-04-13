from constants import *


class UserEntry:
    def __init__(self):
        self.file_location = FILE_LOCATION
        self.file_name = FILE_NAME
        self.file_suffix = FILE_SUFFIX
        self.file_name_save = FILE_NAME_SAVE

        self.component_index = COMPONENT_INDEX_DEFAULT
        self.quantity_index = QUANTITY_INDEX_DEFAULT
        self.n_rows_to_peak = N_ROWS_TO_PEAK_DEFAULT
