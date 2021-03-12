import pandas as pd
# my id generator
import id_generator as id
# Regex library
import re
# system library
import sys

folders = {}
cover_letters = {}
photographies = {}
persons = {}
ending_first_part = 6556


def create_folder(fol_id, fol_name):
    folder = {
        "id": fol_id,
        "name": fol_name
    }

    folders[fol_id] = folder


def create_cover_letter(co_le_id, co_pa, fol_id):
    cover_letter = {
        "id": co_le_id,
        "page number": co_pa,
        "folder id": fol_id
    }

    cover_letters[co_le_id] = cover_letter


def create_person(per_id, f_name, turk_l_name, amr_l_name, husb_name, fath_name, moth_name, gr_fath_name, photo_id):
    person = {
        "id": per_id,
        "first name": f_name,
        "turkish last name": turk_l_name,
        "armenian last name": amr_l_name,
        "husband name": husb_name,
        "father name": fath_name,
        "mother name": moth_name,
        "grand father name": gr_fath_name,
        "photo id": photo_id
    }

    persons[per_id] = person


def create_photograph(pho_id, cop_sen):
    photograph = {
        "id": pho_id,
        "copies sent": cop_sen
    }

    photographies[pho_id] = photograph


def get_df_folder():
    folder_id_val = []
    folder_name_val = []

    for f in folders.values():
        folder_id_val.append(f["id"])
        folder_name_val.append(f["name"])

    if len(folder_id_val) != len(folder_name_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({'ID': folder_id_val, 'Name': folder_name_val})


def get_df_cover_letter():
    cov_let_id_val = []
    cov_let_page_val = []
    cov_let_folder_val = []

    for cl in cover_letters.values():
        cov_let_id_val.append(cl["id"])
        cov_let_page_val.append(cl["page number"])
        cov_let_folder_val.append(cl["folder id"])

    if len(cov_let_id_val) != len(cov_let_page_val) or len(cov_let_page_val) != len(cov_let_folder_val):
        print("FAIL - Cover letter property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({'ID': cov_let_id_val, 'Page': cov_let_page_val, 'Folder ID': cov_let_folder_val})


def get_df_photograph():
    photo_id_val = []
    photo_copy_val = []

    for p in photographies.values():
        photo_id_val.append(p["id"])
        photo_copy_val.append(p["copies sent"])

    if len(photo_id_val) != len(photo_copy_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({'ID': photo_id_val, 'Copies of photograph sent ': photo_copy_val})


def get_df_person():
    per_id_val = []
    per_fn_val = []
    per_tu_ln_val = []
    per_ar_ln_val = []
    per_hn_val = []
    per_fa_n_val = []
    per_mo_n_val = []
    per_gfa_n_val = []
    per_photo_val = []

    for p in persons.values():
        per_id_val.append(p["id"])
        per_fn_val.append(p["first name"])
        per_tu_ln_val.append(p["turkish last name"])
        per_ar_ln_val.append(p["armenian last name"])
        per_hn_val.append(p["husband name"])
        per_fa_n_val.append(p["father name"])
        per_mo_n_val.append(p["mother name"])
        per_gfa_n_val.append(p["grand father name"])
        per_photo_val.append(p["photo id"])

    if len(per_id_val) != len(per_fn_val) or \
            len(per_id_val) != len(per_tu_ln_val) or \
            len(per_id_val) != len(per_ar_ln_val) or \
            len(per_id_val) != len(per_hn_val) or \
            len(per_id_val) != len(per_fa_n_val) or \
            len(per_id_val) != len(per_mo_n_val) or \
            len(per_id_val) != len(per_gfa_n_val) or \
            len(per_id_val) != len(per_photo_val):
        print("FAIL - person property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': per_id_val,
        'First Name': per_fn_val,
        'Turkish Last Name': per_tu_ln_val,
        'Armenian Last Name': per_ar_ln_val,
        'Husband\'s Name': per_hn_val,
        'Father\'s Name': per_fa_n_val,
        'Mother\'s Name': per_mo_n_val,
        'Grandfather\'s Name': per_gfa_n_val,
        'Photograph ID': per_photo_val
    })


def start():
    full_data = pd.read_excel("00_input_data/pou_data.xlsx", sheet_name="Tab A")
    # range of rows (first part)
    df = full_data.iloc[0:ending_first_part]
    # print number of rows
    print("Total rows first Part: {0}".format(len(df)))

    last_folder_id = None
    last_photo_id = None
    # Iterates through the rows
    for index, row in df.iterrows():
        # Checks if cell in column A is not nan (= has value)
        if not pd.isna(row[0]):

            last_folder_id = id.generate_id(row[0])
            if last_folder_id not in folders:
                create_folder(last_folder_id, row[0])

        if not pd.isna(row[2]):
            cover_letter_id = id.generate_random_id()
            create_cover_letter(cover_letter_id, row[2], last_folder_id)

        if not pd.isna(row[6]):
            last_photo_id = id.generate_random_id()
            create_photograph(last_photo_id, row[6])

        if not pd.isna(row[11]):
            person_id = id.generate_random_id()
            turk_last_name = None
            arm_last_name = None
            husband_name = None
            fathers_name = None
            mothers_name = None
            grand_fathers_name = None
            if not pd.isna(row[12]):
                turk_last_name = row[12]
            if not pd.isna(row[13]):
                arm_last_name = row[13]
            if not pd.isna(row[14]):
                husband_name = row[14]
            if not pd.isna(row[15]):
                fathers_name = row[15]
            if not pd.isna(row[16]):
                mothers_name = row[16]
            if not pd.isna(row[17]):
                grand_fathers_name = row[17]
            create_person(person_id, row[11], turk_last_name, arm_last_name, husband_name, fathers_name, mothers_name, grand_fathers_name, last_photo_id)

        # if not pd.isna(row[6]) and pd.isna(row[11]) and pd.isna(row[12]):
        #     print("No person on photograph: ({0})".format(index + 2))

    # -----------------------------------

    df_folder = get_df_folder()
    df_cover_letter = get_df_cover_letter()
    df_photograph = get_df_photograph()
    df_person = get_df_person()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('04_output_data/result.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_folder.to_excel(writer, sheet_name='Folder')
    df_cover_letter.to_excel(writer, sheet_name='Cover Letter')
    df_photograph.to_excel(writer, sheet_name='Photograph')
    df_person.to_excel(writer, sheet_name='Person')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


def second_part():
    full_data = pd.read_excel("00_input_data/pou_data.xlsx", sheet_name="Tab A")
    # range of rows (first part)
    df = full_data.iloc[0:ending_first_part]
    # print number of rows
    print("Total rows first Part: {0}".format(len(df)))

    # Iterates through the rows
    for index, row in df.iterrows():
        # Checks if cell is not nan (= has value)
        if not pd.isna(row[0]):

            # Creates new folder object
            folder = {
                "id": id.generate_id(row[0])
            }

            dh_name = re.search("(DH_(.+))(_\d\d\d)", row[0])
            hr_name = re.search("(HR_(.+))(_\d\d\d)", row[0])
            folder_name = ""

            if dh_name:
                folder["name"] = dh_name.group(1)
            elif hr_name:
                folder["name"] = hr_name.group(1)
            else:
                folder["name"] = row[0]

            folders.append(folder)

    folder_keys = [*folders[0].keys()]

    folder_id_values = []
    folder_name_values = []
    for f in folders:
        folder_id_values.append(f["id"])
        folder_name_values.append(f["name"])

    # print(folder_id_values, folder_name_values)

    # Create a Pandas dataframe from the data.
    df2 = pd.DataFrame({'ID': folder_id_values, 'Name': folder_name_values})

    # Convert the dataframe to an XlsxWriter Excel object.
    df2.to_excel("04_output_data/result.xlsx", sheet_name='Folder')
