__author__ = 'mcharbit'

from xml.etree.cElementTree import SubElement, parse
from collections import OrderedDict
from random import randint
import os
import shutil
import zipfile
import time

class xlsx_gen(object):
    """
    Class to generate simple xlsx file from a template.
    Template must contain a sheet named sheet1 and this sheet must contain at least 1 value (that will be deleted)
    Usage :
        xlsx = xlsx_gen(file_in="Template.xlsx", file_out="Generated_file.xlsx")
        xlsx.write("Quizz title", "A", "1")
        xlsx.write("Quizz date", "A", "2")
        xlsx.write_to_file()
    """
    base_ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

    def __init__(self, file_in, file_out):
        self.file_in = file_in
        self.file_out = file_out
        self.dict = {}
        self.strings_tree = None
        unique_suffix = "_" + str(time.time()) + "_" + str(randint(0, 100))
        self.tmp_zip_file = "zip_tmp" + unique_suffix + ".zip"
        self.tmp_dir = "dir_tmp" + unique_suffix
        self.tmp_sheet_xml_file = self.tmp_dir + "/xl/worksheets/sheet1.xml"
        self.tmp_strings_xml_file = self.tmp_dir + "/xl/sharedStrings.xml"

        if not zipfile.is_zipfile(file_in):
            raise

    def process_shared(self):

        self.strings_tree = parse(self.tmp_strings_xml_file)

        # Removing file from the template. XML structure has been loaded to memory, file will be recreated in the zip
        # archive when calling the write function
        os.remove(self.tmp_strings_xml_file)

        sst = self.strings_tree.getroot()

        if len(self.dict) > 0:

            uniqueCount_file = sst.get("uniqueCount")
            if uniqueCount_file is None:
                uniqueCount = 0
            else:
                uniqueCount = int(uniqueCount_file)

            text_content = []
            ts = sst.findall(".//" + self.base_ns + "t")
            for index, t_tmp in enumerate(ts):
                text_content.append(t_tmp.text)

            for key, value in self.dict.iteritems():

                text_index = None

                # Iterating through the cell values. If the value is already defined, return its index, rather than
                # appending a new si element
                for index, text in enumerate(text_content):
                    if value[0] == text:
                        text_index = index
                        break

                # Value not defined in the file, create a new si Element
                if text_index is None:
                    si = SubElement(sst, self.base_ns + "si")
                    t = SubElement(si, self.base_ns + "t")
                    t.text = value[0]
                    text_index = uniqueCount
                    uniqueCount += 1

                    text_content.append(value[0])

                # Update the text index to be referenced in sheet1.xml
                self.dict[key][1] = text_index

            sst.set("uniqueCount", str(uniqueCount))


    def finalize_shared(self, count_update):

        # Update count attribute of the sst Element with the actual number of cells.
        # Excel will consider broken file otherwise
        if self.strings_tree is not None:
            sst = self.strings_tree.getroot()
            sst.set("count", str(count_update))

            # Write XML tree to file
            self.strings_tree.write(self.tmp_strings_xml_file)
            # Write file to archive
            self.zout.write(self.tmp_strings_xml_file, arcname="xl/sharedStrings.xml")
        else:
            raise

    @staticmethod
    def column_to_index(column):
        total_index=0
        for index, char in enumerate(reversed(column)):
            total_index+= pow(26, index) * (ord(char.lower()) - 96)
        return total_index

    @staticmethod
    def index_to_column(index):
        local_index=0
        column=""

        while index > 0:
            division = index // pow(26, local_index)
            if division > 0:
                local_index+=1

            else:
                if local_index == 0:
                    column += chr((index % pow(26, local_index))+96).upper()
                    index = 0
                else:
                    column += chr((index // pow(26, local_index - 1))+96).upper()
                    index = index % pow(26, local_index - 1)
                    local_index = 0

        return column

    def process_sheet1(self):

        tree = parse(self.tmp_sheet_xml_file)

        # Removing file from the template. XML structure has been loaded to memory, file will be recreated in the zip
        # archive when calling the write function
        os.remove(self.tmp_sheet_xml_file)

        worksheet = tree.getroot()
        sheetData = worksheet.find("./" + self.base_ns + "sheetData")

        # Removing existing cell records from template if any
        file_rows = sheetData.findall("./" + self.base_ns + "row")
        if len(file_rows) > 0:
            for file_row in file_rows:
                sheetData.remove(file_row)

        curent_row_number = None
        curent_cell_ref = None
        total_cell_number = 0
        row = None
        v = None

        for key, value in self.dict.iteritems():

            row_number = value[4]

            # New row ? If so, append a new row Element
            if row_number != curent_row_number:
                # Append the new row
                curent_row_number = row_number
                row = SubElement(sheetData, self.base_ns + "row")
                row.set("r", row_number)

            # Determine if the cell is already defined in the file. If so (multiple writing to the same cell), last
            # writing to the cell (last item of the for loop) will actually be effective
            if key != curent_cell_ref:

                # Append the new cell
                curent_cell_ref = key
                total_cell_number += 1

                c = SubElement(row, self.base_ns + "c")
                style = value[2]
                if style is not None:
                    c.set("s", str(style))
                c.set("r", key)
                c.set("t", "s")
                v = SubElement(c, self.base_ns + "v")

            v.text = str(value[1])

        # Updating sharedstring "count" attribute from the sharedstrings.xml file with the new number of cells
        # Excel will complain otherwise
        self.finalize_shared(total_cell_number)

        # Updating dimension element. Optional, only allows Excel/Open Office to optimize
        # the horizontal / vertical scroll bar size. Not really time consuming though
        min_row = self.dict.values()[0][4]
        max_row = self.dict.values()[-1][4]

        tmp_set = set()
        for cell_data in self.dict.values():
            tmp_set.add(cell_data[3])
        sorted_columns = sorted(tmp_set)

        min_column = sorted_columns[0]
        max_column = sorted_columns[-1]

        new_dim = min_column + str(min_row) + ":" + max_column + str(max_row)

        dimension = worksheet.find("./" + self.base_ns + "dimension")
        if dimension is not None:
            dimension.set("ref", new_dim)

        # Adjusting columns width for the columns that are used
        col_records = None
        cols = worksheet.find("./" + self.base_ns + "cols")
        col = None
        if cols is not None:
            col_records = cols.findall("./" + self.base_ns + "col")
        else:
            cols = SubElement(worksheet, self.base_ns + "cols")

        if col_records is not None:
            # Removing existing col elements if any, keeping only the last one to update it
            for index, col_record in enumerate(col_records):
                if index < len(col_records)-1:
                    cols.remove(col_record)
                else:
                    col = col_record

        if col is None:
            col = SubElement(cols, self.base_ns + "col")
        col.set("min", str(self.column_to_index(min_column)))
        col.set("max", str(self.column_to_index(max_column)))
        col.set("width", "30")
        col.set("customWidth", "1")

        # Writing XML tree to file
        tree.write(self.tmp_sheet_xml_file)
        # Writing file to archive
        self.zout.write(self.tmp_sheet_xml_file, arcname="xl/worksheets/sheet1.xml")


    def write(self, value, column, row, style=None):

        """
        The function is accepting 4 arguments
         * value to be written. Ex : "My text"
         * column and row to be written. Ex : "A", "3"
         * style to be used for the cell. Ex : 2
         There are predefined styles in the current template. If styles need to be redefined, procedure used is :
         - Update the style of a cell in an empty file.
         - Determine the index used for the style that has been created ("s" attribute of the <c> element) in sheet1.xml
         - Edit manually sheet1.xml to delete the corresponding <c> elements (files appears empty, but styles exist)
         - Reference the appropriate style when calling the write function
        """

        # Dict of lists
        # Format of elements : dict[cell_reference] = [cell_text, index_of_cell_text_in_shared_string, style, col, row]
        cell = str(column) + str(row)
        self.dict[cell]=[value, None, style, str(column), str(row)]

    def write_to_file(self):

        """
        Function to be called when all writing is done.
            - Unzips the template archive
            - Calls function to process the sharedstrings.xml and sheet1.xml files according to cells written
            - Leaves other files untouched
            - Creates output archive and return it as a file object or real file depending on the given output parameter
        """

        # We will iterate over the cells of the dictionary and append them to the sheet1.xml file. It is important that
        # cells are correctly sorted by row, then by column. Excel wont display the data if the order is broken
        tmp_dict=OrderedDict(sorted(self.dict.iteritems(), key=lambda t: (int(t[1][4]), t[1][3])))
        self.dict = tmp_dict

        zin = zipfile.ZipFile(self.file_in, mode="r")
        zin.extractall(self.tmp_dir)

        # Check if we were passed a file-like object for output file
        if isinstance(self.file_out, basestring):
            filePassed = False

            # If not, delete output archive if it exists
            if os.path.exists(self.tmp_zip_file):
                os.remove(self.tmp_zip_file)

            self.zout = zipfile.ZipFile(self.tmp_zip_file, mode="w")
        else:
            filePassed = True
            self.zout = zipfile.ZipFile(self.file_out, mode="w")

        ordered_file_list = list(zin.infolist())
        # Processing xl/sharedStrings.xml will return the position of text to be used when
        # processing xl/worksheets/sheet1.xml
        # We want to make sure that one file is processed before the other
        ordered_file_list.sort(key=lambda item: item.filename)

        for item in ordered_file_list:
            if item.filename == 'xl/sharedStrings.xml':
                self.process_shared()
            elif item.filename == 'xl/worksheets/sheet1.xml':
                self.process_sheet1()
            elif item.filename[-1] != '~' :
                # Filtering tmp files from archive.
                # These tmp files are created by the OS when working on the members. Excel will complain if those files
                # are present in the output archive
                self.zout.write(self.tmp_dir + "/" + item.filename, item.filename)

        zin.close()

        if not filePassed:
            self.zout.close()
            shutil.move(self.tmp_zip_file, self.file_out)

        if os.path.exists(self.tmp_dir):
            shutil.rmtree(self.tmp_dir)
