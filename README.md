# gcsv2mmdb
This utility converts Google contacts csv file (outlook version) to Motorola PhoneBook Manager MDB file.

## Prerequisites
* Python 2.x (x86 - In order to work with the Jet driver).
* win32com package for Python

## Usage Instructions

* Export your Google contacts (use the outlook version).
* Convert the contacts csv file to Motorola's mdb file using the script: gCSV2mMDB.py
* Import your contacts to the device using Motorola PhoneBook Manager.

## Common Issues

* If you encounter "Class is not licensed for use" while running the script, please apply the 'Jet 3.5 License fix.reg' from the misc directory.
