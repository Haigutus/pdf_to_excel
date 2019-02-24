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


def move_file(file_path, relative_destination_folder, new_file_name):

    working_directory = os.getcwd()

    file_name = os.path.basename(file_path)
    destination_folder = os.path.join(working_directory, relative_destination_folder)


    if os.path.exists(destination_folder) == False:

        print("Alert - Following path was not found, creating new one -> {}".format(destination_folder))

        os.makedirs(destination_folder)



    destination_path = os.path.join(destination_folder, new_file_name)

    os.rename(file_path, destination_path)

    return destination_path



def list_of_files(path,file_extension):

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

    print matches
    return matches



def parse_assesment_to_excel(assessment_path, database_path):

    utc_now = datetime.utcnow()

    data_dictionary = OrderedDict({"Processed_UTC":utc_now.isoformat()}) # Lets make an dictionary where to keep all the parsed values, lets add time when parsing was started


    assessment_file = open(assessment_path, 'rb')

    parser = PDFParser(assessment_file)
    doc = PDFDocument(parser)
    fields = resolve1(doc.catalog['AcroForm'])['Fields']
    for i in fields:
        field = resolve1(i)
        key, value = field.get('T'), field.get('V')

        print '{}: {} -> {}'.format(key, value, type(value)) # DEBUG

        if type(value) == str:

            unicode_value = unicode(value.decode("iso-8859-1").replace(u"\xfe\xff\x00", u"").replace(u"\x00", u"")) # Lets convert the string to uncide and replace is needed to remove some funny characters
            data_dictionary[key] = [unicode_value]

        elif value == None:
            data_dictionary[key] = [u"Ei"]

        else:
            data_dictionary[key] = [value.name]

    assessment_file.close()


    data_frame = pandas.DataFrame(data_dictionary)
    print data_frame # DEBUG


    if os.path.exists(database_path) == True:

        print "Info - Database file {} already exists, loading previous records".format(database_path)
        exsiting_data = pandas.read_excel(database_path)
        print exsiting_data
        data_frame = exsiting_data.append(data_frame)

        #Create bacup of current database

        move_file(database_path, "database_bacup", "{:%Y%m%d-%H%M%S}_{}".format(utc_now, database_path))




    data_frame.reset_index(drop=True, inplace=True)

    writer = pandas.ExcelWriter(database_path)
    data_frame.to_excel(writer,'Hindamised')
    writer.save()


    return data_dictionary








# Settings
database_path = "hindamiste_andmebaas.xlsx"
incomig_files_path = "incoming"

# PROCESS START

incoming_assessmnets = list_of_files(incomig_files_path, "pdf")

if len(incoming_assessmnets) == 0:
    print "No assessments found - process stop"


for assessment_path in incoming_assessmnets:

    file_name = "{:%Y%m%d-%H%M%S}_{}.pdf".format(datetime.utcnow(), uuid.uuid4()) # Rename to unique filename

    assessment_path = move_file(assessment_path, "unknown", file_name)

    data_dict = parse_assesment_to_excel(assessment_path, database_path)

    move_file(assessment_path, "processed", file_name)