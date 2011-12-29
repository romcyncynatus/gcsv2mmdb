"""
This class implements information gathering from google csv files.
You may get contacts using this class.
The class treat csv files as read only
as it not meant to alter them at all.
"""
import csv
import os
from Contact import Contact, EmptyContactError

class gCSV:
    def __init__(self,gcsv_file_path,override_gcsv_file=False):
        # Set gcsv file path as class property 
        self.gcsv_file_path = gcsv_file_path

        # Check whether gcsv file already exists
        if os.path.exists(self.gcsv_file_path):
            # If asked to override, delete file and create a new gcsv file
            if override_gcsv_file:
                os.remove(self.gcsv_file_path)
                try:
                    gcsv_file = open(gcsv_file_path,"r+b")
                except Exception, e:
                    print "Could not create gCSV file: %s" % gcsv_file_path
                    raise e
                self._InitGcsv(gcsv_file,True)
            # Just open existing gcsv file
            else:
                try:
                    gcsv_file = open(gcsv_file_path,"r+b")
                except Exception, e:
                    print "Could not open gCSV file: %s" % gcsv_file_path
                    raise e
                self._InitGcsv(gcsv_file,False)
        # gcsv doesn't exists. Create a new one.
        else:
            try:
                gcsv_file = open(gcsv_file_path,"r+b")
            except Exception, e:
                print "Could not create gCSV file: %s" % gcsv_file_path
                raise e
            self._InitGcsv(gcsv_file,True)
    
    def _InitGcsv(self,gcsv_file,is_new_gcsv):
        self.gcsv_file = gcsv_file
        
    def GetContacts(self):
        try:
            self.gcsvreader = csv.DictReader(self.gcsv_file)
        except Exception, e:
            print "Could not initialize gcsv reader."
            raise e
        
        gcsvcontacts = [contact for contact in self.gcsvreader]
        contacts = []
        
        # Create contact class for each contact
        for contact in gcsvcontacts:
            # Initialize contact parameters
            Name = ""
            Home=None
            Work1=None
            Work2=None
            Mobile=None
            Email=None
            Email2=None
            Fax=None
            
            # Generate contact name
            if contact["First Name"] != "":
                Name = "%s %s" % (Name, contact["First Name"])
                
            if contact["Middle Name"] != "":
                Name = "%s %s" % (Name, contact["Middle Name"])
                
            if contact["Last Name"] != "":
                Name = "%s %s" % (Name, contact["Last Name"])
                
            if contact["Company"] != "":
                Name = "%s %s" % (Name, contact["Company"])
                
            # Remove leading space char
            Name = Name[1:]
                
            # Generate rest of contact's data
            if contact["Home Phone"] != "":
                Home = contact["Home Phone"]
                
            if contact["Business Phone"] != "":
                Work1 = contact["Business Phone"]
                
            if contact["Business Phone 2"] != "":
                Work2 = contact["Business Phone 2"]
                
            if contact["Mobile Phone"] != "":
                Mobile = contact["Mobile Phone"]
                
            if contact["E-mail Address"] != "":
                Email = contact["E-mail Address"]
                
            if contact["E-mail 2 Address"] != "":
                Email2 = contact["E-mail 2 Address"]
                
            if contact["Home Fax"] != "":
                Fax = contact["Home Fax"]
            
            try:
                contacts.append(Contact(Name,Home,Mobile,Work1,Work2,Email,Email2,Fax))
            except EmptyContactError:
                # We just wont add empty contacts
                pass
                
            
        return contacts 
            
  
    def Close(self):
        self.gcsv_file.close()