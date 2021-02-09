# Sync korskyrkan register with contacts

## requirements
Python 2.7 or greater (only tested in python 3)

# For linux
[mdbtools](https://github.com/mdbtools/mdbtools)

# For windows
https://www.microsoft.com/en-us/download/confirmation.aspx?id=13255
and also run
`pip install pyodbc`

## How it works

Open the `sync_settings.json` file and change the `file_path` to wherever your `.mdb` file is located.

It will then convert the mdb to a csv which can be imported into google contacts.

