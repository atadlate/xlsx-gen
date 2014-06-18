__author__ = 'mcharbit'

from xml.etree.cElementTree import SubElement, parse
from generator import XlsxGen
import os
import time
from io import StringIO
import codecs
import zipfile
import shutil

def extract_all_xlsx():
    for dir_tuple in os.walk('.'):
        for file in dir_tuple[2]:
            file_name_wo_ext , file_ext = os.path.splitext(file)
            if file_ext == '.xlsx':
                file_in = os.path.join(dir_tuple[0], file)
                dir_out = os.path.join(dir_tuple[0], file_name_wo_ext)

                if os.path.exists(dir_out):
                    shutil.rmtree(dir_out)

                zin = zipfile.ZipFile(file_in, mode="r")
                zin.extractall(dir_out)

def visual_parser(file, file_log):

    def parser(element, level):

        file_log.write("\t"*level + element.tag + " / " + str(element.text) + " / " + str(element.items()) + "\n")
        if len(element) != 0:
            level+=1
            for node in element:
                parser(node, level)

    local_tree = parse(file)
    root = local_tree.getroot()
    file_log.write("Parsing " + file + "\n")
    parser(root, 0)
    file_log.write("\n")

def print_xml(parsed_dir, logfile=False):

    if not logfile:
        file_log = StringIO()
        method = "print"
    else:
        file_log = codecs.open(os.path.join(parsed_dir, "structure.log"), encoding='utf-8', mode='w')
        method = "file"

    for root, dirs, files in os.walk(parsed_dir):
        for file in files:

            if file.endswith(".xml") or file.endswith(".rels") :
                visual_parser(os.path.join(root, file), file_log)

    if method == "print":
        print(file_log.getvalue())

    file_log.close()


def demo(nb_execution=1):

    try:
        nb_exec = int(nb_execution)
    except:
        nb_exec = 1

    mean_time = 0

    for i in range(nb_exec):

        # For performance profiling
        start = time.time()

        xlsx = XlsxGen(file_in="Template.xlsx", file_out="generated_file.xlsx")

        xlsx.write("Quizz title", "A", "1", 2)
        xlsx.write("Quizz date", "A", "2", 1)
        xlsx.write("Room name", "A", "3", 1)
        xlsx.write("Common Core Tags:", "A", "5", 2)
        xlsx.write("Tag1", "B", "5", 1)
        xlsx.write("Tag2", "C", "5", 1)
        xlsx.write("Tag3", "D", "5", 1)
        xlsx.write("Student Names", "A", "7", 4)
        xlsx.write("Total Score (0 - 100)", "B", "7", 4)
        xlsx.write("Number of correct answers", "C", "7", 4)
        xlsx.write("Question 1", "D", "7", 4)
        xlsx.write("Question 2", "E", "7", 4)
        xlsx.write(u"Johny Strangename " + chr(630) + " " + chr(631), "A", "8", 1)
        xlsx.write("50", "B", "8", 3)
        xlsx.write("1", "C", "8", 3)
        xlsx.write("My ass", "D", "8", 5)
        xlsx.write("looks good", "E", "8", 6)
        xlsx.write(u"Robert from Sweden " + chr(510) + " " + chr(571), "A", "9", 1)
        xlsx.write("100", "B", "9", 3)
        xlsx.write("2", "C", "9", 3)
        xlsx.write("My hair", "D", "9", 6)
        xlsx.write("look good", "E", "9", 6)
        xlsx.write("Class scoring", "A", "10", 7)
        xlsx.write("50%", "B", "10", 8)
        xlsx.write("50%", "C", "10", 8)
        xlsx.write("0%", "D", "10", 8)
        xlsx.write("100%", "E", "10", 8)

        for row in range(12, 60):
            for column_index in range(26):
                column = chr(column_index+97).upper()
                xlsx.write("Dummy text", column, str(row))

        xlsx.write_to_file()

        stop = time.time()
        mean_time += (stop - start)

    mean_time /= nb_exec
    print("Mean time elapsed: {}s".format(mean_time))

# Calling demo
demo()

# Extracting content of the xlsx file to a subdirectory
extract_all_xlsx()

# Printing a visual representation of the arborescence of all xml files from the xlsx archive.
#   - Logfile to True = generate to File
#   - Logfile to False = print to screen
print_xml("generated_file", logfile=True)
