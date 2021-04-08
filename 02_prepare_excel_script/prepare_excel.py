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
persons = {}
# input & output file
input_excel_file = "00_input_data/POU import file.xlsx"
input_tab = "Sheet1"
starting_first_part = 0
ending_first_part = 6558
starting_second_part = 6558
ending_second_part = 6970
output_path = "04_output_data/"


def create_folder(fol_id, fol_name):
    folder = {
        "id": fol_id,
        "name": fol_name,
        "sent": None,
        "cover letter id": [],
        "photograph id": []
    }

    folders[fol_id] = folder


def update_folder(fold_id, co_le_id):
    if not folders[fold_id]:
        print("FAIL - Folder ID invalid")
        raise SystemExit(0)

    if co_le_id:
        folders[fold_id]["cover letter id"].append(co_le_id)


def create_cover_letter(co_le_id, co_pa, addressor, addressee, date_greg, date_hicri, photo_police, ministry, rel_det,
                        mat_beo, mat_amkt):
    cover_letter = {
        "id": co_le_id,
        "page number": co_pa,
        "addressor": addressor,
        "addressee": addressee,
        "date greg": date_greg,
        "date hicri": date_hicri,
        "photographer police": photo_police,
        "ministry o foreign affairs": ministry,
        "relevant details": rel_det,
        "matching beo": mat_beo,
        "matching a mkt": mat_amkt,
        "photograph id": []
    }

    cover_letters[co_le_id] = cover_letter


def update_cover_letter(co_le_id, co_pa, addressor, addressee, date_greg, date_hicri, photo_police, ministry, rel_det,
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

    if date_greg:
        cover_letters[co_le_id]["date greg"] = date_greg

    if date_hicri:
        cover_letters[co_le_id]["date hicri"] = date_hicri

    if photo_police:
        cover_letters[co_le_id]["photographer police"] = photo_police

    if ministry:
        cover_letters[co_le_id]["ministry o foreign affairs"] = ministry

    if rel_det:
        cover_letters[co_le_id]["relevant details"] = rel_det

    if mat_beo:
        cover_letters[co_le_id]["matching beo"] = mat_beo

    if mat_amkt:
        cover_letters[co_le_id]["matching a mkt"] = mat_amkt

    if photo_id:
        cover_letters[co_le_id]["photograph id"].append(photo_id)


def create_person(per_id, gen, f_name, turk_l_name, arm_l_name, husb_name, fath_name, moth_name,
                  gr_fath_name, kin_rel, house, des_coun, des_city, na_app, prof, reli, eye, compl,
                          mouth, hair, mu, beard, face, height):
    person = {
        "id": per_id,
        "gender": gen,
        "first name": f_name,
        "turkish last name": turk_l_name,
        "armenian last name": arm_l_name,
        "husband name": husb_name,
        "father name": fath_name,
        "mother name": moth_name,
        "grand father name": gr_fath_name,
        "kin relationship": kin_rel,
        "house": house,
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
        "height": height
    }

    persons[per_id] = person


def create_photograph(pho_id, cop_sen, leffen, firar):
    photograph = {
        "id": pho_id,
        "same": None,
        "leffen": leffen,
        "wording": None,
        "copies sent": cop_sen,
        "firar": firar,
        "leave family": None,
        "anchor": None,
        "pass info": None,
        "pass info celb": None,
        "pass info varak": None,
        "date pass": None,
        "people": None,
        "physical copy id": []
    }

    photographs[pho_id] = photograph


def get_df_folder():
    folder_id_val = []
    folder_name_val = []
    folder_sent_ist_val = []
    folder_cov_let_val = []
    folder_phot_val = []

    for f in folders.values():
        folder_id_val.append(f["id"])
        folder_name_val.append(f["name"])
        folder_sent_ist_val.append(f["sent"])
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
        'Prints enclosed and sent to Istanbul': folder_sent_ist_val,
        'Cover Letter ID\'s': folder_cov_let_val,
        'Photograph ID\'s': folder_phot_val
    })


def get_df_cover_letter():
    cov_let_id_val = []
    cov_let_page_val = []
    cov_let_addressor_val = []
    cov_let_addressee_val = []
    cov_let_date_greg_val = []
    cov_let_date_hicri_val = []
    cov_let_pho_pol_val = []
    cov_let_minis_val = []
    cov_let_rel_val = []
    cov_let_mat_beo = []
    cov_let_mat_amkt = []
    cov_let_photo_val = []

    for cl in cover_letters.values():
        cov_let_id_val.append(cl["id"])
        cov_let_page_val.append(cl["page number"])
        cov_let_addressor_val.append(cl["addressor"])
        cov_let_addressee_val.append(cl["addressee"])
        cov_let_date_greg_val.append(cl["date greg"])
        cov_let_date_hicri_val.append(cl["date hicri"])
        cov_let_pho_pol_val.append(cl["photographer police"])
        cov_let_minis_val.append(cl["ministry o foreign affairs"])
        cov_let_rel_val.append(cl["relevant details"])
        cov_let_mat_beo.append(cl["matching beo"])
        cov_let_mat_amkt.append(cl["matching a mkt"])
        cov_let_photo_val.append(";".join(cl["photograph id"]))

    if len(cov_let_id_val) != len(cov_let_page_val) or \
            len(cov_let_id_val) != len(cov_let_addressor_val) or \
            len(cov_let_id_val) != len(cov_let_addressee_val) or \
            len(cov_let_id_val) != len(cov_let_date_greg_val) or \
            len(cov_let_id_val) != len(cov_let_date_hicri_val) or \
            len(cov_let_id_val) != len(cov_let_pho_pol_val) or \
            len(cov_let_id_val) != len(cov_let_minis_val) or \
            len(cov_let_id_val) != len(cov_let_rel_val) or \
            len(cov_let_id_val) != len(cov_let_mat_beo) or \
            len(cov_let_id_val) != len(cov_let_mat_amkt) or \
            len(cov_let_id_val) != len(cov_let_photo_val):
        print("FAIL - Cover letter property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': cov_let_id_val,
        'Page': cov_let_page_val,
        'Addressor': cov_let_addressor_val,
        'Addressee': cov_let_addressee_val,
        'Date Gregorian': cov_let_date_greg_val,
        'Date Hicri': cov_let_date_hicri_val,
        'Studio Photographer or Police': cov_let_pho_pol_val,
        'Ministry of Foreign Affairs': cov_let_minis_val,
        'Relevant Details': cov_let_rel_val,
        'Matching File in BEO': cov_let_mat_beo,
        'Matching File in A MKT': cov_let_mat_amkt,
        'Photograph ID': cov_let_photo_val
    })


def get_df_photograph():
    photo_id_val = []
    photo_same_val = []
    photo_leffen_val = []
    photo_wording_val = []
    photo_copy_val = []
    photo_firar_val = []
    photo_leave_val = []
    photo_anchor_val = []
    photo_pass_info_val = []
    photo_pass_info_celb_val = []
    photo_pass_info_varak_val = []
    photo_date_pass_val = []
    photo_people_val = []
    photo_physical_co_val = []

    for p in photographs.values():
        photo_id_val.append(p["id"])
        photo_same_val.append(p["same"])
        photo_leffen_val.append(p["leffen"])
        photo_wording_val.append(p["wording"])
        photo_copy_val.append(p["copies sent"])
        photo_firar_val.append(p["firar"])
        photo_leave_val.append(p["leave family"])
        photo_anchor_val.append(p["anchor"])
        photo_pass_info_val.append(p["pass info"])
        photo_pass_info_celb_val.append(p["pass info celb"])
        photo_pass_info_varak_val.append(p["pass info varak"])
        photo_date_pass_val.append(p["date pass"])
        photo_people_val.append(p["people"])
        photo_physical_co_val.append(";".join(p["physical copy id"]))

    if len(photo_id_val) != len(photo_copy_val):
        print("FAIL - Folder property values not same length")
        raise SystemExit(0)

    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': photo_id_val,
        'Same as': photo_same_val,
        'Leffen': photo_leffen_val,
        'Wording regarding photography': photo_wording_val,
        'Copies of photograph sent ': photo_copy_val,
        'Firar-I iade': photo_firar_val,
        'Did they leave their family': photo_leave_val,
        'Anchoring individual': photo_anchor_val,
        'Passport information': photo_pass_info_val,
        'Passport information (Celb)': photo_pass_info_celb_val,
        'Passport information (Varak)': photo_pass_info_varak_val,
        'Date of Passport': photo_date_pass_val,
        'People on Picture': photo_people_val,
        'Physical Copy ID': photo_physical_co_val
    })


def get_df_physical_copy():
    # Create a Pandas dataframe from the data.
    return pd.DataFrame({
        'ID': [],
        'Seal of State': [],
        'Seal of State Issuer': [],
        'Second Seal': [],
        'Second Seal Issuer': [],
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
    per_kin_val = []
    per_hou_val = []
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
        per_kin_val.append(p["kin relationship"])
        per_hou_val.append(p["house"])
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

    if len(per_id_val) != len(per_gen_val) or \
            len(per_id_val) != len(per_fn_val) or \
            len(per_id_val) != len(per_tu_ln_val) or \
            len(per_id_val) != len(per_ar_ln_val) or \
            len(per_id_val) != len(per_hn_val) or \
            len(per_id_val) != len(per_fa_n_val) or \
            len(per_id_val) != len(per_mo_n_val) or \
            len(per_id_val) != len(per_gfa_n_val) or \
            len(per_id_val) != len(per_kin_val) or \
            len(per_id_val) != len(per_hou_val) or \
            len(per_id_val) != len(per_des_co_val) or \
            len(per_id_val) != len(per_des_ci_val) or \
            len(per_id_val) != len(per_na_app_val) or \
            len(per_id_val) != len(per_prof_val) or \
            len(per_id_val) != len(per_rel_val) or \
            len(per_id_val) != len(per_eye_val) or \
            len(per_id_val) != len(per_com_val) or \
            len(per_id_val) != len(per_mou_val) or \
            len(per_id_val) != len(per_hair_val) or \
            len(per_id_val) != len(per_mus_val) or \
            len(per_id_val) != len(per_bea_val) or \
            len(per_id_val) != len(per_fac_val) or \
            len(per_id_val) != len(per_hei_val):
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
        'Kin Relationship': per_kin_val,
        'House': per_hou_val,
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
        'Height': per_hei_val
    })


def start():
    full_data = pd.read_excel(input_excel_file, sheet_name=input_tab)

    # --------------- FIRST PART - START --------------------
    # range of rows (first part)
    df = full_data.iloc[starting_first_part:ending_first_part]
    # print number of rows
    print("Total rows first Part: {0}".format(len(df)))

    last_folder_id = None
    first_co_le_id = None
    first_co_le_with_page = None
    last_cover_letter_id = None

    # Iterates through the rows
    for index, row in df.iterrows():

        # Evaluates the properties for cover letter
        addressor = None if pd.isna(row[4]) else row[4]
        addressee = None if pd.isna(row[5]) else row[5]
        date_hicri = None if pd.isna(row[6]) else row[6]
        date_greg = None if pd.isna(row[7]) else row[7]
        rel_detail = None if pd.isna(row[39]) else row[39]
        ministry_fa = None if pd.isna(row[49]) else row[49]
        match_beo = None if pd.isna(row[53]) else row[53]
        match_a_mkt = None if pd.isna(row[54]) else row[54]
        photo_police = None if pd.isna(row[98]) else row[98]

        # print(addressor, addressee, rel_detail, match_beo, match_a_mkt)

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
            create_cover_letter(first_co_le_id, page_number, addressor, addressee, date_greg, date_hicri,
                                photo_police, ministry_fa, rel_detail,
                                match_beo, match_a_mkt)
            update_folder(last_folder_id, first_co_le_id)
        else:
            # Checks if there is a page number (= cover letter starts)
            if not pd.isna(row[3]):
                # Checks if first cover letter was created without a page number
                if not first_co_le_with_page:
                    update_cover_letter(last_cover_letter_id, row[3], addressor, addressee, date_greg, date_hicri,
                                        photo_police, ministry_fa, rel_detail,
                                        match_beo, match_a_mkt, None)
                    first_co_le_with_page = True
                else:
                    last_cover_letter_id = id.generate_random_id()
                    create_cover_letter(last_cover_letter_id, row[3], addressor, addressee, date_greg, date_hicri,
                                        photo_police, ministry_fa, rel_detail,
                                        match_beo, match_a_mkt)
                    update_folder(last_folder_id, last_cover_letter_id)
            else:
                update_cover_letter(last_cover_letter_id, None, addressor, addressee, date_greg, date_hicri,
                                    photo_police, ministry_fa, rel_detail,
                                    match_beo, match_a_mkt, None)

        if not pd.isna(row[11]):
            # Property for photograph
            leffen = "False" if pd.isna(row[8]) else "True"
            firar = "False" if pd.isna(row[12]) else "True"

            last_photo_id = id.generate_random_id()
            create_photograph(last_photo_id, row[11], leffen, firar)
            update_cover_letter(last_cover_letter_id, None, None, None, None, None, None, None, None, None,
                                None, last_photo_id)

        # Checks if there is a first name
        if not pd.isna(row[19]):
            person_id = id.generate_random_id()
            gender = None if pd.isna(row[15]) else row[15]
            turk_last_name = None if pd.isna(row[20]) else row[20]
            arm_last_name = None if pd.isna(row[21]) else row[21]
            husband_name = None if pd.isna(row[22]) else row[22]
            fathers_name = None if pd.isna(row[23]) else row[23]
            mothers_name = None if pd.isna(row[24]) else row[24]
            grand_fathers_name = None if pd.isna(row[25]) else row[25]
            kin_relation = None if pd.isna(row[26]) else row[26]
            house = None if pd.isna(row[31]) else row[31]
            destination_country = None if pd.isna(row[32]) else row[32]
            destination_city = None if pd.isna(row[33]) else row[33]
            name_appear = None if pd.isna(row[56]) else row[56]
            profession = None if pd.isna(row[79]) else row[79]
            religion = None if pd.isna(row[80]) else row[80]
            eye_color = None if pd.isna(row[81]) else row[81]
            complexion = None if pd.isna(row[82]) else row[82]
            mouth_nose = None if pd.isna(row[83]) else row[83]
            hair_color = None if pd.isna(row[84]) else row[84]
            mustache = None if pd.isna(row[85]) else row[85]
            beard = None if pd.isna(row[86]) else row[86]
            face = None if pd.isna(row[87]) else row[87]
            height = None if pd.isna(row[88]) else row[88]

            create_person(person_id, gender, row[19], turk_last_name, arm_last_name, husband_name, fathers_name,
                          mothers_name, grand_fathers_name, kin_relation, house, destination_city, destination_country,
                          name_appear, profession, religion, eye_color, complexion, mouth_nose, hair_color, mustache,
                          beard, face, height)

        # --------------- TESTING CODE --------------------
        # Test: Folder name occurs at least twice
        # if not pd.isna(row[1]):
        #     last_folder_id = id.generate_id(row[1])
        #     if last_folder_id in folders:
        #         print("Folder name in {0} appeared before".format(index + 2))

        # Test: Cover Letter without Photograph
        # if not pd.isna(row[3]) and pd.isna(row[7]):
        #     print("Cover Letter without Photograph", index + 2)

        # Test: Folder name and a new Photo in the same line
        # if not pd.isna(row[1]) and not pd.isna(row[7]):
        #     print("Folder name and new Photo", index + 2)

        # Test: Photograph without person
        # if not pd.isna(row[7]) and pd.isna(row[12]) and pd.isna(row[13]):
        #     print("No person on photograph: ({0})".format(index + 2), row[7], row[12], row[13])

        # Test: Person without first name but has Turkish last name
    # --------------- FIRST PART - END --------------------

    df_folder = get_df_folder()
    df_cover_letter = get_df_cover_letter()
    df_physical_copy = get_df_physical_copy()
    df_photograph = get_df_photograph()
    df_person = get_df_person()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer_folder = pd.ExcelWriter(output_path + "folder.xlsx", engine='xlsxwriter')
    writer_cover_letter = pd.ExcelWriter(output_path + "cover_letter.xlsx", engine='xlsxwriter')
    writer_physical_copy = pd.ExcelWriter(output_path + "physical_copy.xlsx", engine='xlsxwriter')
    writer_photograph = pd.ExcelWriter(output_path + "photograph.xlsx", engine='xlsxwriter')
    writer_person = pd.ExcelWriter(output_path + "person.xlsx", engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_folder.to_excel(writer_folder, sheet_name='Folder')
    df_cover_letter.to_excel(writer_cover_letter, sheet_name='Cover Letter')
    df_physical_copy.to_excel(writer_physical_copy, sheet_name='Physical Copy')
    df_photograph.to_excel(writer_photograph, sheet_name='Photograph')
    df_person.to_excel(writer_person, sheet_name='Person')

    # Close the Pandas Excel writer and output the Excel file.
    writer_folder.save()
    writer_cover_letter.save()
    writer_physical_copy.save()
    writer_photograph.save()
    writer_person.save()
