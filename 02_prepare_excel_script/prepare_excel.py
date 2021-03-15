import pandas as pd
# my id generator
import id_generator as id
# Regex library
import re
# system library
import sys

folders = {}
cover_letters = {}
photographs = {}
persons = {}
# input & output file
input_excel_file = "00_input_data/output.xlsx"
input_tab = "Sheet1"
starting_first_part = 0
ending_first_part = 6556
output_excel_file = "04_output_data/result.xlsx"


def create_folder(fol_id, fol_name):
    folder = {
        "id": fol_id,
        "name": fol_name
    }

    folders[fol_id] = folder


def create_cover_letter(co_le_id, co_pa, fol_id, addressor, addressee, date, police, ministry, spec_com, rel_det):
    cover_letter = {
        "id": co_le_id,
        "page number": co_pa,
        "folder id": fol_id,
        "addressor": addressor,
        "addressee": addressee,
        "date": date,
        "police department": police,
        "ministry o foreign affairs": ministry,
        "special commission": spec_com,
        "relevant details": rel_det,
    }

    cover_letters[co_le_id] = cover_letter


def update_cover_letter(co_le_id, co_pa, addressor, addressee, date, police, ministry, spec_com, rel_det):
    if not cover_letters[co_le_id]:
        print("FAIL - Cover Letter ID invalid")
        raise SystemExit(0)

    if co_pa:
        cover_letters[co_le_id]["page number"] = co_pa

    if addressor:
        cover_letters[co_le_id]["addressor"] = addressor

    if addressee:
        cover_letters[co_le_id]["addressee"] = addressee

    if date:
        cover_letters[co_le_id]["date"] = date

    if police:
        cover_letters[co_le_id]["police department"] = police

    if ministry:
        cover_letters[co_le_id]["ministry o foreign affairs"] = ministry

    if spec_com:
        cover_letters[co_le_id]["special commission"] = spec_com

    if rel_det:
        cover_letters[co_le_id]["relevant details"] = rel_det


def create_person(per_id, f_name, turk_l_name, amr_l_name, husb_name, fath_name, moth_name,
                  gr_fath_name, bi_place, or_town, or_kaza, des_coun, des_city, photo_id):
    person = {
        "id": per_id,
        "first name": f_name,
        "turkish last name": turk_l_name,
        "armenian last name": amr_l_name,
        "husband name": husb_name,
        "father name": fath_name,
        "mother name": moth_name,
        "grand father name": gr_fath_name,
        "birth place": bi_place,
        "origin town": or_town,
        "origin kaza": or_kaza,
        "destination country": des_coun,
        "destination city": des_city,
        "photo id": photo_id
    }

    persons[per_id] = person


def create_photograph(pho_id, cop_sen, co_le_id):
    photograph = {
        "id": pho_id,
        "copies sent": cop_sen,
        "cover letter id": co_le_id
    }

    photographs[pho_id] = photograph


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
    return pd.DataFrame({
        'ID': folder_id_val,
        'Name': folder_name_val
    })


def get_df_cover_letter():
    cov_let_id_val = []
    cov_let_page_val = []
    cov_let_folder_val = []
    cov_let_addressor_val = []
    cov_let_addressee_val = []
    cov_let_date_val = []
    cov_let_pol_val = []
    cov_let_minis_val = []
    cov_let_spec_val = []
    cov_let_rel_val = []

    for cl in cover_letters.values():
        cov_let_id_val.append(cl["id"])
        cov_let_page_val.append(cl["page number"])
        cov_let_folder_val.append(cl["folder id"])
        cov_let_addressor_val.append(cl["addressor"])
        cov_let_addressee_val.append(cl["addressee"])
        cov_let_date_val.append(cl["date"])
        cov_let_pol_val.append(cl["police department"])
        cov_let_minis_val.append(cl["ministry o foreign affairs"])
        cov_let_spec_val.append(cl["special commission"])
        cov_let_rel_val.append(cl["relevant details"])

    if len(cov_let_id_val) != len(cov_let_page_val) or \
            len(cov_let_id_val) != len(cov_let_folder_val) or \
            len(cov_let_id_val) != len(cov_let_addressor_val) or \
            len(cov_let_id_val) != len(cov_let_addressee_val) or \
            len(cov_let_id_val) != len(cov_let_date_val) or \
            len(cov_let_id_val) != len(cov_let_pol_val) or \
            len(cov_let_id_val) != len(cov_let_minis_val) or \
            len(cov_let_id_val) != len(cov_let_spec_val) or \
            len(cov_let_id_val) != len(cov_let_rel_val):
        print("FAIL - Cover letter property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': cov_let_id_val,
        'Page': cov_let_page_val,
        'Folder ID': cov_let_folder_val,
        'Addressor': cov_let_addressor_val,
        'Addressee': cov_let_addressee_val,
        'Date': cov_let_date_val,
        'Police Department': cov_let_pol_val,
        'Ministry of Foreign Affairs': cov_let_minis_val,
        'Special Commission': cov_let_spec_val,
        'Relevant Details': cov_let_rel_val
    })


def get_df_photograph():
    photo_id_val = []
    photo_copy_val = []
    photo_cov_let_val = []

    for p in photographs.values():
        photo_id_val.append(p["id"])
        photo_copy_val.append(p["copies sent"])
        photo_cov_let_val.append(p["cover letter id"])

    if len(photo_id_val) != len(photo_copy_val) or \
            len(photo_id_val) != len(photo_cov_let_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': photo_id_val,
        'Copies of photograph sent ': photo_copy_val,
        'Cover Letter ID': photo_cov_let_val
    })


def get_df_person():
    per_id_val = []
    per_fn_val = []
    per_tu_ln_val = []
    per_ar_ln_val = []
    per_hn_val = []
    per_fa_n_val = []
    per_mo_n_val = []
    per_gfa_n_val = []
    per_birth_val = []
    per_or_to_val = []
    per_or_kaz_val = []
    per_des_co_val = []
    per_des_ci_val = []
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
        per_birth_val.append(p["birth place"])
        per_or_to_val.append(p["origin town"])
        per_or_kaz_val.append(p["origin kaza"])
        per_des_co_val.append(p["destination country"])
        per_des_ci_val.append(p["destination city"])
        per_photo_val.append(p["photo id"])

    if len(per_id_val) != len(per_fn_val) or \
            len(per_id_val) != len(per_tu_ln_val) or \
            len(per_id_val) != len(per_ar_ln_val) or \
            len(per_id_val) != len(per_hn_val) or \
            len(per_id_val) != len(per_fa_n_val) or \
            len(per_id_val) != len(per_mo_n_val) or \
            len(per_id_val) != len(per_gfa_n_val) or \
            len(per_id_val) != len(per_birth_val) or \
            len(per_id_val) != len(per_or_to_val) or \
            len(per_id_val) != len(per_or_kaz_val) or \
            len(per_id_val) != len(per_des_co_val) or \
            len(per_id_val) != len(per_des_ci_val) or \
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
        'Birth Place': per_birth_val,
        'Origin Town': per_or_to_val,
        'Origin Kaza': per_or_kaz_val,
        'Destination Country': per_des_co_val,
        'Destination City': per_des_ci_val,
        'Photograph ID': per_photo_val
    })


def start():
    full_data = pd.read_excel(input_excel_file, sheet_name=input_tab)
    # range of rows (first part)
    df = full_data.iloc[starting_first_part:ending_first_part]
    # print number of rows
    print("Total rows first Part: {0}".format(len(df)))

    last_folder_id = None
    first_co_le_id = None
    first_co_le_with_page = None
    last_cover_letter_id = None
    last_photo_id = None
    # Iterates through the rows
    for index, row in df.iterrows():
        # Checks if cell in column B is not nan (= has folder name)
        if not pd.isna(row[1]):
            last_folder_id = id.generate_id(row[1])
            # Creates new folder if name does not exist
            if last_folder_id not in folders:
                create_folder(last_folder_id, row[1])

            # Sometimes when there is a folder name the page number occurs in the same line.
            # In that case it must be indicated. So when later in a row without a folder name and a visible page number
            # it is clear if the addressor/addressee/date information must be added to a new cover letter or must be
            # added to the first cover letter.
            if pd.isna(row[3]):
                page_number = None
                first_co_le_with_page = False
            else:
                page_number = row[3]
                first_co_le_with_page = True

            addressor = None if pd.isna(row[29]) else row[29]
            addressee = None if pd.isna(row[30]) else row[30]
            date = None if pd.isna(row[32]) else row[32]
            police_department = not pd.isna(row[33])
            ministry_fa = not pd.isna(row[34])
            special_commission = not pd.isna(row[35])
            rel_detail = None if pd.isna(row[36]) else row[36]

            first_co_le_id = id.generate_random_id()
            last_cover_letter_id = first_co_le_id
            create_cover_letter(first_co_le_id, page_number, last_folder_id, addressor, addressee, date,
                                police_department, ministry_fa, special_commission, rel_detail)
        else:
            addressor = None if pd.isna(row[29]) else row[29]
            addressee = None if pd.isna(row[30]) else row[30]
            date = None if pd.isna(row[32]) else row[32]
            police_department = not pd.isna(row[33])
            ministry_fa = not pd.isna(row[34])
            special_commission = not pd.isna(row[35])
            rel_detail = None if pd.isna(row[36]) else row[36]

            # Checks if there is a page number (= cover letter starts)
            if not pd.isna(row[3]):
                if not first_co_le_with_page:
                    update_cover_letter(last_cover_letter_id, row[3], addressor, addressee, date,
                                        police_department, ministry_fa, special_commission, rel_detail)
                else:
                    last_cover_letter_id = id.generate_random_id()
                    create_cover_letter(last_cover_letter_id, row[3], last_folder_id, addressor, addressee, date,
                                        police_department, ministry_fa, special_commission, rel_detail)
            else:
                update_cover_letter(last_cover_letter_id, None, addressor, addressee, date,
                                    police_department, ministry_fa, special_commission, rel_detail)

        # if not pd.isna(row[1]) and not pd.isna(row[29]):
        #     print(row[29], row[1], index + 2)

        if not pd.isna(row[7]):
            last_photo_id = id.generate_random_id()
            create_photograph(last_photo_id, row[7], last_cover_letter_id)

        if not pd.isna(row[12]):
            person_id = id.generate_random_id()
            turk_last_name = None if pd.isna(row[13]) else row[13]
            arm_last_name = None if pd.isna(row[14]) else row[14]
            husband_name = None if pd.isna(row[15]) else row[15]
            fathers_name = None if pd.isna(row[16]) else row[16]
            mothers_name = None if pd.isna(row[17]) else row[17]
            grand_fathers_name = None if pd.isna(row[18]) else row[18]
            birth_place = None if pd.isna(row[23]) else row[23]
            origin_town = None if pd.isna(row[24]) else row[24]
            origin_kaza = None if pd.isna(row[25]) else row[25]
            destination_country = None if pd.isna(row[27]) else row[27]
            destination_city = None if pd.isna(row[28]) else row[28]

            create_person(person_id, row[12], turk_last_name, arm_last_name, husband_name, fathers_name,
                          mothers_name, grand_fathers_name, birth_place, origin_town, origin_kaza,
                          destination_country, destination_city, last_photo_id)

        # Lists all photograph without person
        # if not pd.isna(row[7]) and pd.isna(row[12]) and pd.isna(row[13]):
        #     print("No person on photograph: ({0})".format(index + 2))

    # -----------------------------------

    df_folder = get_df_folder()
    df_cover_letter = get_df_cover_letter()
    df_photograph = get_df_photograph()
    df_person = get_df_person()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_folder.to_excel(writer, sheet_name='Folder')
    df_cover_letter.to_excel(writer, sheet_name='Cover Letter')
    df_photograph.to_excel(writer, sheet_name='Photograph')
    df_person.to_excel(writer, sheet_name='Person')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
