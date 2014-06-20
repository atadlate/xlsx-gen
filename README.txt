Class to generate simple xlsx file from a template.
Template must contain a sheet named sheet1 and this sheet must contain at least 1 value (that will be deleted)
Usage :
        xlsx = xlsx_gen(file_in="Template.xlsx", file_out="Generated_file.xlsx")
        xlsx.write("Quizz title", "A", "1")
        xlsx.write("Quizz date", "A", "2")
        xlsx.write_to_file()

