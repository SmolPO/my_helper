import logging
logging.basicConfig(filename="B:/my_helper/log_file.log", level=logging.INFO)

db_keys = {"auto": "(gov_number, brand, model, track_number, id)",
           "drivers": "(family, name, surname, birthday, passport, adr, id)",
           "workers": "(family, name, surname, post, birthday, passport, passport_got, adr, live_adr, phone, inn, "
                      "snils, "
                      "numb_contract, date_contract, "
                      "numb_h, numb_card_h, date_h, "
                      "numb_study, numb_study_card, d_study, "
                      "numb_protocol, numb_card, d_protocol, object, "
                      "d_vac_1, d_vac_2, place, vac_doc, vac_type, status, id)",
           "itrs": "(family, name, surname, post, passport, passport_got, adr, live_adr, auto, inn, "
                   "snils, n_employment_contract, date_employment_contract, "
                   "ot_protocol, ot_date, ot_card, "
                   "PTM_protocol, PTM_date, PTM_card, "
                   "es_protocol, es_group, es_card, es_date, "
                   "h_protocol, h_date, h_group, h_card, "
                   "industrial_save, "
                   "st_protocol, st_card, st_date, birthday, "
                   " d_vac_1, d_vac_2, place, vac_doc, vac_type, status, id)",
           "contracts": "(name, customer, number, date, object, type_work, place, "
                        "price, date_end, nds, avans, status, id)",
           "company": "(company, adr, ogrn, inn, kpp, bik, korbill, rbill, bank, big_boss, "
                   "big_post, big_at, big_d_at, mng_boss, mng_post, mng_at, mng_d_at, status, id)",
           "bosses": "(family, name, surname, post, email, phone, sex, company, id)",
           "materials": "(name, measure, value, provider, contract, id)",
           "links": "(name, link, id)",
           "finance": "(date, value, recipient, comment, id)",
           "bills": "(date, value, buyer, name_file, comment, id)",
           "tools": "(name, count, codes, date, id)",
           "tenders": "(name, customer, object, work, part, kp_1, out_1, d_out_1, kp_2, out_2, d_out_2, id)"}


def get_person(person):
    row = "SELECT * FROM workers WHERE family={0}".format(person)
    logging.debug(row)
    return row


def get_from_db(fields, db):
    insert = "SELECT " + fields + " FROM " + db
    logging.debug(insert)
    return insert


def add_to_db(data, table):
    result = "('" + "', '".join(data) + "')"
    names_colomn = db_keys[table]
    row = "INSERT INTO {0} {1} VALUES {2}".format(table, names_colomn, result)
    logging.debug(row)
    return row


def my_update(data, table):
    fields = db_keys.get(table)
    my_id = data[-1]
    result = "('" + "', '".join(data) + "')"
    row = "UPDATE {0} SET {1} = {2} where id = '{3}'".format(table, fields, result, my_id)
    logging.debug(row)
    return row


