import psycopg2

conn = psycopg2.connect("host=95.163.249.246 dbname=Vertical_db user=office password=9024EgrGvz#m87Y1")
print("OK conn")
db = conn.cursor()
list_db_ = ["CREATE TABLE contracts (name text, customer text, number text, date text, object text, type_work text, "
            "place text, id text)",
            "CREATE TABLE bosses (family text, name text, surname text, post text, email text, phone text, sex text"
               ", id text)",
            "CREATE TABLE drivers (family text, name text, surname text, birthday text, passport text, adr text,"
            " id text)",
            "CREATE TABLE company (company text, adr text, ogrn text, inn text, kpp text, bik text, korbill text, "
            "rbill text, bank text, family text, name text, surname text, post text, count_attorney text,"
            "date_attorney text, id text)",
            "CREATE TABLE auto (model text, brand text, gov_number text, track_number text, id text)",
            "CREATE TABLE drivers (family text, name text, surname text, birthday text, passport text, adr text,"
            " id text)",
            "CREATE TABLE contracts (name text, customer text, number text, date text, object text, type_work text, "
               "place text, id text)",
            "CREATE TABLE bosses (family text, name text, surname text, post text, email text, phone text, sex text,"
            " id text)",
            "CREATE TABLE links (name text, link text, id text)",
            "CREATE TABLE materials (name text, measure text, value text, provider text, contract text, id text)",
            "CREATE TABLE workers (family text, name text, surname text, birthday text, post text, phone text, "
            "passport text, passport_got text, adr text, live_adr text, inn text, snils text, numb_contract text, "
            "date_contract text, numb_h text, numb_group_h text, date_h text, numb_study text, numb_study_card text,"
            "d_study text, numb_protocol text, numb_card text, d_protocol text, status text, id text)",
            "CREATE TABLE itrs (family text, name text, surname text, post text, passport text, passport_date text, "
            "passport_got text, adr text, live_adr text, auto text, inn text, "
            "snils text, n_employment_contract text, date_employment_contract text, "
            "ot_protocol text, ot_date text, ot_card text, "
            "PTM_protocol text, PTM_date text, PTM_card text, "
            "es_protocol text, es_group text, es_card text, es_date text, "
            "h_protocol text, h_date text, h_group text, h_card text, "
            "industrial_save text, "
            "st_protocol text, st_card text, st_date text, birthday text, status text, id text)",
            "CREATE TABLE finance (id text, date text, value text, recipient text, comment text, id text)",
            "CREATE TABLE bills (date text, value text, buyer text, name_file text, comment text, id text)",
            "CREATE TABLE notes (date text, name text, id text)"]
for item in list_db_:
    try:
        db.execute(item)
        conn.commit()
        print("OK item" + item[13:20])
    except:
        print("ERROR " + item)
print("init db!! Let's go!")
