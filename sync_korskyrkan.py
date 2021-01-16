import csv
import json
import os


table_name = "person"


def read_settings():
    with open('sync_settings.json') as json_file:
        data = json.load(json_file)
        return data


def convert_database_to_csv(file_path):
    cmd = "mdb-export " + file_path + " Person > person.csv"
    os.system(cmd)
    #result = subprocess.run(cmd.split(" "), stdout=subprocess.PIPE)
    #if result.stderr is not None:
    #    raise Exception("Bash command went wrong, " + result.stderr + "," + result.stdout)
    #return result.stdout


def read_values_from_csv_file():
    with open('person.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        result = []
        for row in csv_reader:
            if line_count == 0:
                print('Column names are ' + ", ".join(row))
                line_count += 1
            else:
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
    with open('person_google.csv', mode='w') as person_google_csv:
        writer = csv.DictWriter(person_google_csv, fieldnames=fieldnames, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        writer.writeheader()
        for row in result:
            writer.writerow({'Name': row["fnamn"], 'Family Name': row["enamn"], 'E-mail 1 - Type': '* INTERNET', 'E-mail 1 - Value': row["email"], "Phone 1 - Type": "Mobile", "Phone 1 - Value": row["mobiltelenr"]})


settings = read_settings()
convert_database_to_csv(settings["file_path"])
result = read_values_from_csv_file()
filtered_result = filter_result_without_email(result)
save_as_google_csv(filtered_result)
