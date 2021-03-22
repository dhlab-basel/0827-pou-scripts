import pandas as pd
# my id generator
import id_generator as id
# helper module
import helper_folder as fold
# system library
import sys

folders = {}
cover_letters = {}
photographs = {}
physical_copies = {}
persons = {}

folder_id_start = 2
person_id_start = 2

# input & output file
input_excel_file = "00_input_data/output.xlsx"
input_tab = "Sheet1"
starting_first_part = 0
ending_first_part = 6556
starting_second_part = 6557
ending_second_part = 6961
output_excel_file = "04_output_data/result.xlsx"


def create_folder(fol_id, fol_name):
    global folder_id_start
    folder = {
        "id": folder_id_start,
        "name": fol_name,
        "cover letter id": [],
        "photograph id": []
    }

    folder_id_start += 1

    folders[fol_id] = folder


def update_folder(fold_id, co_le_id):
    if not folders[fold_id]:
        print("FAIL - Folder ID invalid")
        raise SystemExit(0)

    # if co_le_id:
    #     folders[fold_id]["cover letter id"].append(co_le_id)


def create_cover_letter(co_le_id, co_pa, addressor, addressee, date, police, ministry, spec_com, rel_det,
                        mat_beo, mat_amkt):
    cover_letter = {
        "id": co_le_id,
        "page number": co_pa,
        "addressor": addressor,
        "addressee": addressee,
        "date": date,
        "police department": police,
        "ministry o foreign affairs": ministry,
        "special commission": spec_com,
        "relevant details": rel_det,
        "matching beo": mat_beo,
        "matching a mkt": mat_amkt,
        "photograph id": []
    }

    cover_letters[co_le_id] = cover_letter


def update_cover_letter(co_le_id, co_pa, addressor, addressee, date, police, ministry, spec_com, rel_det,
                        mat_beo, mat_amkt, photo_id):
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

    if mat_beo:
        cover_letters[co_le_id]["matching beo"] = mat_beo

    if mat_amkt:
        cover_letters[co_le_id]["matching a mkt"] = mat_amkt

    if photo_id:
        cover_letters[co_le_id]["photograph id"].append(photo_id)


def create_person(per_id, gen, f_name, turk_l_name, amr_l_name, husb_name, fath_name, moth_name,
                  gr_fath_name, bi_place, or_town, or_kaza, des_coun, des_city, na_app, prof, reli, eye, compl,
                          mouth, hair, mu, beard, face, height, house):
    global person_id_start
    person = {
        "id": person_id_start,
        "gender": gen,
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
        "name appear": na_app,
        "profession": prof,
        "religion": reli,
        "eye color": eye,
        "complexion": compl,
        "mouth": mouth,
        "hair color": hair,
        "mustache": mu,
        "beard": beard,
        "face": face,
        "height": height,
        "house": house
    }

    person_id_start += 1

    persons[per_id] = person


def create_photograph(pho_id, cop_sen):
    photograph = {
        "id": pho_id,
        "copies sent": cop_sen
    }

    photographs[pho_id] = photograph


def get_df_folder():
    folder_id_val = []
    folder_name_val = []
    folder_cov_let_val = []
    folder_phot_val = []

    for f in folders.values():
        folder_id_val.append(f["id"])
        folder_name_val.append(f["name"])
        folder_cov_let_val.append(";".join(f["cover letter id"]))
        folder_phot_val.append(";".join(f["photograph id"]))

    if len(folder_id_val) != len(folder_name_val) or \
            len(folder_id_val) != len(folder_cov_let_val) or \
            len(folder_id_val) != len(folder_phot_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': folder_id_val,
        'Name': folder_name_val,
        'Cover Letter ID\'s': folder_cov_let_val,
        'Photograph ID\'s': folder_phot_val
    })


def get_df_cover_letter():
    cov_let_id_val = []
    cov_let_page_val = []
    cov_let_addressor_val = []
    cov_let_addressee_val = []
    cov_let_date_val = []
    cov_let_pol_val = []
    cov_let_minis_val = []
    cov_let_spec_val = []
    cov_let_rel_val = []
    cov_let_mat_beo = []
    cov_let_mat_amkt = []
    cov_let_photo_val = []

    for cl in cover_letters.values():
        cov_let_id_val.append(cl["id"])
        cov_let_page_val.append(cl["page number"])
        cov_let_addressor_val.append(cl["addressor"])
        cov_let_addressee_val.append(cl["addressee"])
        cov_let_date_val.append(cl["date"])
        cov_let_pol_val.append(cl["police department"])
        cov_let_minis_val.append(cl["ministry o foreign affairs"])
        cov_let_spec_val.append(cl["special commission"])
        cov_let_rel_val.append(cl["relevant details"])
        cov_let_mat_beo.append(cl["matching beo"])
        cov_let_mat_amkt.append(cl["matching a mkt"])
        cov_let_photo_val.append(";".join(cl["photograph id"]))

    if len(cov_let_id_val) != len(cov_let_page_val) or \
            len(cov_let_id_val) != len(cov_let_addressor_val) or \
            len(cov_let_id_val) != len(cov_let_addressee_val) or \
            len(cov_let_id_val) != len(cov_let_date_val) or \
            len(cov_let_id_val) != len(cov_let_pol_val) or \
            len(cov_let_id_val) != len(cov_let_minis_val) or \
            len(cov_let_id_val) != len(cov_let_spec_val) or \
            len(cov_let_id_val) != len(cov_let_rel_val) or \
            len(cov_let_id_val) != len(cov_let_mat_beo) or \
            len(cov_let_id_val) != len(cov_let_mat_amkt) or \
            len(cov_let_id_val) != len(cov_let_photo_val):
        print("FAIL - Cover letter property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': [],
        'Page': [],
        'Addressor': [],
        'Addressee': [],
        'Date': [],
        'Police Department': [],
        'Ministry of Foreign Affairs': [],
        'Special Commission': [],
        'Relevant Details': [],
        'Matching File in BEO': [],
        'Matching File in A MKT': [],
        'Photograph ID': []
    })


def get_df_photograph():
    photo_id_val = []
    photo_copy_val = []

    for p in photographs.values():
        photo_id_val.append(p["id"])
        photo_copy_val.append(p["copies sent"])

    if len(photo_id_val) != len(photo_copy_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': [],
        'Same as': [],
        'Leffen': [],
        'Wording regarding photography': [],
        'Copies of photograph sent ': [],
        'Firar-I iade': [],
        'Did they leave their family': [],
        'Anchoring individual': [],
        'Passport information': [],
        'Passport information (Celb)': [],
        'Passport information (Varak)': [],
        'Date of Passport': [],
        'People on Picture': [],
        'Physical Copy ID': []
    })


def get_df_physical_copy():
    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': [],
        'Seal of State': [],
        'Second SealIssuer': [],
        'Bueraucratic Stamp': [],
        'Place of Studio\'s Photographer\'s Name': [],
        'Photographer': [],
        'Location of Photographer': [],
        'Date of Document': [],
        'Date on Photograph': [],
        'Handwritten on front': [],
        'Numbered': [],
        'Perforated': [],
        'Printed information on Front': [],
        'Writing on Front': [],
        'Date of Photograph': [],
        'Color of Ink': [],
        'Other notes': []
    })


def get_df_person():
    per_id_val = []
    per_gen_val = []
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
    per_na_app_val = []
    per_prof_val = []
    per_rel_val = []
    per_eye_val = []
    per_com_val = []
    per_mou_val = []
    per_hair_val = []
    per_mus_val = []
    per_bea_val = []
    per_fac_val = []
    per_hei_val = []
    per_hou_val = []

    for p in persons.values():
        per_id_val.append(p["id"])
        per_gen_val.append(p["gender"])
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
        per_na_app_val.append(p["name appear"])
        per_prof_val.append(p["profession"])
        per_rel_val.append(p["religion"])
        per_eye_val.append(p["eye color"])
        per_com_val.append(p["complexion"])
        per_mou_val.append(p["mouth"])
        per_hair_val.append(p["hair color"])
        per_mus_val.append(p["mustache"])
        per_bea_val.append(p["beard"])
        per_fac_val.append(p["face"])
        per_hei_val.append(p["height"])
        per_hou_val.append(p["house"])

    if len(per_id_val) != len(per_gen_val) or \
            len(per_id_val) != len(per_fn_val) or \
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
            len(per_id_val) != len(per_prof_val) or \
            len(per_id_val) != len(per_rel_val) or \
            len(per_id_val) != len(per_eye_val) or \
            len(per_id_val) != len(per_com_val) or \
            len(per_id_val) != len(per_mou_val) or \
            len(per_id_val) != len(per_hair_val) or \
            len(per_id_val) != len(per_mus_val) or \
            len(per_id_val) != len(per_bea_val) or \
            len(per_id_val) != len(per_fac_val) or \
            len(per_id_val) != len(per_hei_val) or\
            len(per_id_val) != len(per_hou_val):
        print("FAIL - person property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': per_id_val,
        'Gender': per_gen_val,
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
        'Name also appear in': per_na_app_val,
        'Profession': per_prof_val,
        'Religion': per_rel_val,
        'Eye Color': per_eye_val,
        'Complexion': per_com_val,
        'Mouth/Nose': per_mou_val,
        'Hair Color': per_hair_val,
        'Mustache': per_mus_val,
        'Beard': per_bea_val,
        'Face': per_fac_val,
        'Height': per_hei_val,
        'House': per_hou_val
    })


def start():
    full_data = pd.read_excel(input_excel_file, sheet_name=input_tab)

    # range of rows (first part)
    df1 = full_data.iloc[starting_first_part:ending_first_part]
    # range of rows (first part)
    df2 = full_data.iloc[starting_second_part:ending_second_part]

    frames = [df1, df2]
    full_df = pd.concat(frames)
    # print number of rows
    print("Total rows second Part: {0}".format(len(full_df)))

    last_folder_id = None
    first_co_le_id = None
    first_co_le_with_page = None
    last_cover_letter_id = None

    # Iterates through the rows
    for index, row in full_df.iterrows():

        # Evaluates the properties for cover letter
        addressor = None if pd.isna(row[29]) else row[29]
        addressee = None if pd.isna(row[30]) else row[30]
        date = None if pd.isna(row[32]) else row[32]
        police_department = not pd.isna(row[33])
        ministry_fa = not pd.isna(row[34])
        special_commission = not pd.isna(row[35])
        rel_detail = None if pd.isna(row[36]) else row[36]
        match_beo = None if pd.isna(row[52]) else row[52]
        match_a_mkt = None if pd.isna(row[55]) else row[55]

        # Checks if cell in column B is not nan (= has folder name)
        if not pd.isna(row[1]):
            last_folder_id = id.generate_id(row[1])
            # Creates new folder if name does not exist
            if last_folder_id not in folders:
                create_folder(last_folder_id, row[1])

            # Sometimes when there is a folder name, the page number occurs in the same line.
            # In that case it must be indicated. So when later in a row without a folder name and a visible page number
            # it is clear if the addressor/addressee/date information must be added to a new cover letter or must be
            # added to the first cover letter.
            if pd.isna(row[3]):
                page_number = None
                first_co_le_with_page = False
            else:
                page_number = row[3]
                first_co_le_with_page = True

            first_co_le_id = id.generate_random_id()
            last_cover_letter_id = first_co_le_id
            create_cover_letter(first_co_le_id, page_number, addressor, addressee, date,
                                police_department, ministry_fa, special_commission, rel_detail,
                                match_beo, match_a_mkt)
            update_folder(last_folder_id, first_co_le_id)
        else:
            # Checks if there is a page number (= cover letter starts)
            if not pd.isna(row[3]):
                # Checks if first cover letter was created without a page number
                if not first_co_le_with_page:
                    update_cover_letter(last_cover_letter_id, row[3], addressor, addressee, date,
                                        police_department, ministry_fa, special_commission, rel_detail,
                                        match_beo, match_a_mkt, None)
                    first_co_le_with_page = True
                else:
                    last_cover_letter_id = id.generate_random_id()
                    create_cover_letter(last_cover_letter_id, row[3], addressor, addressee, date,
                                        police_department, ministry_fa, special_commission, rel_detail,
                                        match_beo, match_a_mkt)
                    update_folder(last_folder_id, last_cover_letter_id)
            else:
                update_cover_letter(last_cover_letter_id, None, addressor, addressee, date,
                                    police_department, ministry_fa, special_commission, rel_detail,
                                    match_beo, match_a_mkt, None)

        if not pd.isna(row[7]):
            last_photo_id = id.generate_random_id()
            create_photograph(last_photo_id, row[7])
            update_cover_letter(last_cover_letter_id, None, None, None, None, None, None, None, None, None,
                                None, last_photo_id)

        # Checks if there is a first name
        if not pd.isna(row[12]):
            person_id = id.generate_random_id()
            gender = None if pd.isna(row[11]) else row[11]
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
            name_appear = None if pd.isna(row[57]) else row[57]
            profession = None if pd.isna(row[84]) else row[84]
            religion = None if pd.isna(row[85]) else row[85]
            eye_color = None if pd.isna(row[86]) else row[86]
            complexion = None if pd.isna(row[87]) else row[87]
            mouth_nose = None if pd.isna(row[88]) else row[88]
            hair_color = None if pd.isna(row[89]) else row[89]
            mustache = None if pd.isna(row[90]) else row[90]
            beard = None if pd.isna(row[91]) else row[91]
            face = None if pd.isna(row[92]) else row[92]
            height = None if pd.isna(row[93]) else row[93]
            house = None if pd.isna(row[99]) else row[99]

            create_person(person_id, gender, row[12], turk_last_name, arm_last_name, husband_name, fathers_name,
                          mothers_name, grand_fathers_name, birth_place, origin_town, origin_kaza,
                          destination_country, destination_city, name_appear, profession, religion, eye_color, complexion,
                          mouth_nose, hair_color, mustache, beard, face, height, house)

    # --------------- SECOND PART - START --------------------
    # # range of rows (first part)
    # df2 = full_data.iloc[starting_second_part:ending_second_part]
    # # print number of rows
    # print("Total rows second Part: {0}".format(len(df2)))
    #
    # # Iterates through the rows
    # for index, row in df2.iterrows():
    #
    #     # Checks if cell in column B is not nan (= has folder name)
    #     if not pd.isna(row[1]):
    #         fold_obj = fold.get_name(row[1])
    #         fold_id = id.generate_id(fold_obj["name"])
    #
    #         # Creates new folder if name does not exist
    #         if fold_id not in folders:
    #             create_folder(fold_id, fold_obj["name"])
    # --------------- SECOND PART - END --------------------

    df_folder = get_df_folder()
    df_cover_letter = get_df_cover_letter()
    df_photograph = get_df_photograph()
    df_physical_copy = get_df_physical_copy()
    df_person = get_df_person()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_folder.to_excel(writer, sheet_name='Folder', index=False)
    df_cover_letter.to_excel(writer, sheet_name='Cover Letter', index=False)
    df_photograph.to_excel(writer, sheet_name='Photograph', index=False)
    df_physical_copy.to_excel(writer, sheet_name='Physical Copy', index=False)
    df_person.to_excel(writer, sheet_name='Person', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
