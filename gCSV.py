"""
This class implements information gathering from google csv files.
You may get contacts using this class.
The class treat csv files as read only
as it not meant to alter them at all.
"""
import csv

class gCSV:
    def __init__(self,csv_file_path):
        try:
            self.csvfile = open(csv_file_path,"rb")
        except Exception, e:
            print "Could not open CSV file: %s" % csv_file_path
            raise e
        
    def GetContacts(self):
        try:
            self.gcsvreader = csv.DictReader(self.csvfile)
        except Exception, e:
            print "Could not initialize csv reader."
            raise e
        
        return [contact for contact in self.gcsvreader]
        
    def Close(self):
        self.csvfile.close()