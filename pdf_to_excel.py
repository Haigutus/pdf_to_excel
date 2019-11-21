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


    data_frame = pandas.DataFrame(data_dictionary)

    # It seems some PDF editors add hidden fields after filling out, lets limit to allowed fields
    allowed_columns = ['Processed_UTC', 'Elukoht', 'Nr', 'Teostaja', 'Teost_kuup', 'Amet', 'Tootukassas', 'Psyhholoog', 'Psyhhiaater_raviarst', 'Ode_sots', 'Elamisting', 'Oma', 'Yyritud', 'Kellegi_pool', 'Sots_pind', 'Varjupaigas', 'Turvakodus', 'Abivajadus', 'Maj_toimetulek', 'Maj_likert', 'Volad_kohust', 'Elab_elukaaslasega', 'Suhted_elukaaslane', 'Elukaasl_vanus', 'Elukaasl_sugu', 'Kooselu', 'Elab_lastega', 'Suhted_lapsed', 'Laste_arv', 'Lapsed_info', 'Elab_emaga', 'Suhted_ema', 'vanem1_vanus', 'Elab_sugulastega', 'Suhted_sugulased', 'Sugulased_info', 'Elab_sopradega', 'Suhted_sobrad', 'Sobrad_info', 'Elab_tuttavatega', 'Suhted_muud', 'Muud_info', 'Var_haigused', 'Ravimid', 'Ravimi_info', 'Psyhaired', 'Psyhaired_info', 'Var_alkoravi', 'Alkoravi_info', 'Meeleolu', 'DEP', 'UAR', 'PAF', 'SAR', 'AST', 'INS', 'AUDIT', 'CIWA', 'MoCa', 'Tarb_planeeritud', 'Tarb_impuls_reakt', 'Tarb_pidev', 'Tyypkogused', 'Tyypyhikud', 'Tsyklid', 'Tsyklid_info', 'Okserefleks', 'Funktsioon', 'Pos_mojurid', 'Eesmark', 'Kuritarv_raskus', 'Ravi_mot', 'Psyhhosots_sekk', 'Psyhhosots_tyyp', 'Psyhhiaatr_ravi', 'vanem1_sugu', 'Tookoht', 'Suunaja', 'Haridus', 'Nimi', 'Joomapaevad', 'SMS', 'Perearst', 'Synniaasta', 'Koned', 'Kindlustus', 'Sugu', 'Elab_isaga', 'Suhted_isa', 'vanem1_vanus_2', 'vanem1_sugu_2', 'Peaparandus', 'Tung', 'Kained_paevad', 'Jooma_aeg', 'Kaine_aeg', 'F1', 'F2', 'F3', 'F4', 'GHQ', 'Haigused_info', 'Tung_info', 'Lahedastega', 'Juhututtavatega', 'Tooandjast_soltuv', 'Tervisekaebused', 'Tervis_info', 'Mot_info', 'Lab_kuup', 'GGT', 'ALAT', 'ASAT', 'CDT']
    data_frame = data_frame[allowed_columns]

    if debug:
        print list(data_frame.columns) # DEBUG


    if os.path.exists(database_path) == True:

        print "Info  - Database file {} already exists, loading previous records".format(database_path)
        existing_data = pandas.read_excel(database_path)

        if debug:
            print existing_data

        data_frame = existing_data.append(data_frame, sort=False)

        # Create backup of current database
        move_file(database_path, "database_backup", "{:%Y%m%d-%H}_{}".format(utc_now, uuid.uuid4())) # %Y%m%d-%H%M%S to keep every version, right now all bacups on same hour will be written over.




    data_frame.reset_index(drop=True, inplace=True)

    writer = pandas.ExcelWriter(database_path, engine='xlsxwriter')
    data_frame[allowed_columns].to_excel(writer,'Hindamised', encoding='utf8')
    writer.save()

    return data_dictionary



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