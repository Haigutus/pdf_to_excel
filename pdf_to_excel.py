import sys
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

import pandas
from datetime import datetime
from collections import OrderedDict
import os.path
import uuid


pandas.set_option('display.max_rows', 10)
pandas.set_option('display.max_columns', 20)

debug = False


def move_file(file_path, relative_destination_folder, new_file_name):

    working_directory = os.getcwd()

    file_name = os.path.basename(file_path)
    destination_folder = os.path.join(working_directory, relative_destination_folder)


    if os.path.exists(destination_folder) == False:

        print("Alert - Following path was not found, creating new one -> {}".format(destination_folder))

        os.makedirs(destination_folder)



    destination_path = os.path.join(destination_folder, new_file_name)

    if debug:
        print("Moving file from -> to")
        print(file_path, destination_path)

    os.rename(file_path, destination_path)

    return destination_path



def list_of_files(path, file_extension):

    if os.path.exists(path) == False:

        print("Error - path does not exist -> {}".format(path))

    matches = []
    for filename in os.listdir(path):



        if filename.endswith(file_extension):
            #logging.info("Processing file:"+filename)
            matches.append(path + "//" + filename)
        else:
            print "Not a {} file: {}".format(file_extension, filename)
            #logging.warning("Not a {} file: {}".format(file_extension,file_text[0]))

    if debug:
        print matches

    return matches



def parse_assessment_to_excel(assessment_path, database_path):

    utc_now = datetime.utcnow()

    data_dictionary = OrderedDict({"Processed_UTC":utc_now.isoformat()}) # Lets make a dictionary where all the parsed values are kept, lets add time when parsing was started
    # TODO add also processed file name


    assessment_file = open(assessment_path, 'rb')

    parser = PDFParser(assessment_file)
    doc = PDFDocument(parser)
    fields = resolve1(doc.catalog['AcroForm'])['Fields']
    for i in fields:
        field = resolve1(i)
        key, value = field.get('T'), field.get('V')

        if debug:
            print '{}: {} -> {}'.format(key, value, type(value)) # DEBUG

        if type(value) == str:

            unicode_value = unicode(value.decode("iso-8859-1").replace(u"\xfe\xff\x00", u"").replace(u"\x00", u"").replace(u'\xfe\xff', u"")) # Lets convert the string to unicode and replace is needed to remove some funny characters
            data_dictionary[key] = [unicode_value]

        elif value == None:
            data_dictionary[key] = [u"ei"]

        else:
            data_dictionary[key] = [value.name]

            if value.name == "Off":
                data_dictionary[key] = [u"ei"]

            if value.name == "Yes":
                data_dictionary[key] = [u"jah"]

    assessment_file.close()

    # Create pandas dataframe for exporting data
    data_frame = pandas.DataFrame(data_dictionary)

    if debug:
        print list(data_frame.columns) # DEBUG


    if os.path.exists(database_path) == True:

        print "Info  - Database file {} already exists, loading previous records".format(database_path)
        existing_data = pandas.read_excel(database_path, index_col=0) # TODO set first column as index

        if debug:
            print existing_data

        # Add to exsiting data
        data_frame = existing_data.append(data_frame, sort=False)

        # Fix index numbering
        data_frame = data_frame.reset_index(drop=True) # Fix index numbering

        # Create backup of current database
        move_file(database_path, "database_backup", "{:%Y%m%dT%H%M%S}_{}".format(utc_now, uuid.uuid4())) # Create unique filename for each bacup


    # Export to excel and add formatting

    sheet_name = "Hindamised"

    writer = pandas.ExcelWriter(database_path, engine='xlsxwriter')
    data_frame.to_excel(writer, sheet_name, encoding='utf8')

    # Get sheet to do some formatting
    sheet = writer.sheets[sheet_name]

    # Set default column size, if this does not work you are missing XslxWriter module
    first_col = 1
    last_col  = len(data_frame.columns)
    width     = 25
    sheet.set_column(first_col, last_col, width)

    # freeze column names and ID column
    sheet.freeze_panes(1, 1)

    # Apply filter to excel
    first_row = 0
    last_row = len(data_frame)
    sheet.autofilter(first_row, first_col, last_row, last_col)

    # Save the file
    writer.save()

    return data_dictionary

# Run only when the script itself is ran, if script is imported by another script then this part will not run
if __name__ == '__main__':
    # SETTINGS
    database_path        = "HindamisteAndmebaas.xlsx"
    incoming_files_path  = "incoming"
    processed_files_path = "processed"
    unknown_files_path   = "unknown"

    # PROCESS START

    incoming_assessments = list_of_files(incoming_files_path, "pdf")

    if len(incoming_assessments) == 0:
        print "No assessments found - process stop"


    for assessment_path in incoming_assessments:

        print("\nInfo  - processing {}".format(assessment_path))

        # Rename to unique filename as we are keeping them on filesystem
        file_name = "{:%Y%m%d-%H%M%S}_{}.pdf".format(datetime.utcnow(), uuid.uuid4())

        # Move file to unknown status, if it is erronous it will stay there but wont stop processing of new incoming files
        assessment_path = move_file(assessment_path, unknown_files_path, file_name)

        try:
            # Parse file
            data_dict = parse_assessment_to_excel(assessment_path, database_path)

        except Exception as error:

                print("Error - processing failed")
                print(error.__class__.__name__)
                print(error)
                continue

        # Move file to processed status
        move_file(assessment_path, processed_files_path, file_name)

        print("Info  - processing done")