import csv
import json
import os
import pyodbc


WINDOWS = True

def convert_database_to_csv_windows(file_path):
    table_name = "Person"

    # set up some constants
    MDB = file_path
    DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    PWD = 'pw'

    # connect to db
    con = pyodbc.connect('DRIVER={};DBQ={};'.format(DRV,MDB))
    cur = con.cursor()

    # run a query and get the results 
    SQL = 'SELECT * FROM ' + table_name + ';' # your query goes here
    rows = cur.execute(SQL).fetchall()
    cur.close()
    con.close()

    # you could change the mode from 'w' to 'a' (append) for any subsequent queries
    with open('person.csv', 'w') as fou:
        csv_writer = csv.writer(fou) # default field-delimiter is ","
        csv_writer.writerows(rows)


def read_settings():
    with open('sync_settings.json') as json_file:
        data = json.load(json_file)
        return data


def convert_database_to_csv(file_path):
    cmd = "mdb-export " + file_path + " Person > person.csv"
    os.system(cmd)


def read_values_from_csv_file(isWindows):
    with open('person.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        result = []
        for row in csv_reader:
            if line_count == 0 and not isWindows:
                print('Column names are ' + ", ".join(row))
                line_count += 1
            if len(row) == 0:
                continue
                
            result.append({"fnamn": row[1], "enamn": row[2], "telenr": row[16], "mobiltelenr": row[36], "email": row[30]})

            line_count += 1
        print('Processed ' + str(line_count) + ' lines.')
        return result


def filter_result_without_email(result):
    filtered_result = []
    for row in result:
        if(row["email"] != ""):
            filtered_result.append(row)

    return filtered_result


def save_as_google_csv(result):
    fieldnames = "Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value".split(",")
    with open('person_google.csv', mode='w', newline='') as person_google_csv:
        writer = csv.DictWriter(person_google_csv, fieldnames=fieldnames, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        writer.writeheader()
        for row in result:
            writer.writerow({'Name': row["fnamn"], 'Family Name': row["enamn"], 'E-mail 1 - Type': '* INTERNET', 'E-mail 1 - Value': row["email"], "Phone 1 - Type": "Mobile", "Phone 1 - Value": row["mobiltelenr"]})


def save_unique_emails(result):
    emails = {}
    unique_emails = []
    for row in result:
        if row["email"] not in emails:
            emails[row["email"]] = True
            unique_emails.append(row["email"])

    with open("person_unique_email.csv", "w") as text_file:
        text_file.write("\n".join(unique_emails))



settings = read_settings()
convert_database_to_csv_windows(settings["file_path"]) if WINDOWS else convert_database_to_csv(settings["file_path"])
result = read_values_from_csv_file(WINDOWS)
filtered_result = filter_result_without_email(result)
save_as_google_csv(filtered_result)
save_unique_emails(filtered_result)
