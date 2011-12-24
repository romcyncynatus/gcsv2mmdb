from gCSV import gCSV
from mMDB import mMDB
import sys

ARGS_COUNT = 3

def print_usage():
    print "Usage: %s <input csv file> <output mdb file>" % sys.argv[0]

def is_ascii(s):
    return all(ord(c) < 128 for c in s)
       
def convert_gcsv_to_mmdb(gcsv,mmdb):
    contacts = gcsv.GetContacts()
    for contact in contacts:
        # Initialize contact parameters
        ContactName = ""
        Language = "English"
        Home=None
        Work1=None
        Work2=None
        Mobile=None
        Email=None
        Fax=None
        
        # Generate contact name
        if contact["First Name"] != "":
            ContactName = "%s %s" % (ContactName, contact["First Name"])
            
        if contact["Middle Name"] != "":
            ContactName = "%s %s" % (ContactName, contact["Middle Name"])
            
        if contact["Last Name"] != "":
            ContactName = "%s %s" % (ContactName, contact["Last Name"])
            
        if contact["Company"] != "":
            ContactName = "%s %s" % (ContactName, contact["Company"])
            
        # Remove leading space char
        ContactName = ContactName[1:]
            
        # Determine contact name language
        # (Badly implemented. Another idea is required)
        if not is_ascii(ContactName):
           Language = "Hebrew"
            
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
            
        if contact["Home Fax"] != "":
            Fax = contact["Home Fax"]
  
        # Add contact to mmdb file.
        # Only real data will be added. None parameters will be discarded.
        mmdb.AddContact(ContactName,Language,Home,Work1,Work2,Mobile,Email,Fax)     

def main():
    # Validate program arguments. Print usage and terminate upon incorrect usage.
    if len(sys.argv) != ARGS_COUNT:
        print_usage()
        sys.exit(1)
    
    # Create data objects and convert gcsv to mmdb
    try:
        # Create google csv file instance using the input file.
        gcsv = gCSV(sys.argv[1])
        
        # Create motorola mdb file instance for the output file.
        mmdb = mMDB(sys.argv[2])
        
        # Convert gcsv to mmdb
        convert_gcsv_to_mmdb(gcsv,mmdb)
        
        # All went fine, let the world know :)
        print "Motorola mdb file was created successfully."
        print "Import your contacts using Motorola PhoneBook Manager."
    
    except Exception, e:
        print "An error occurred while converting Google CSV file to Motorola mdb file!"
        print "Details: %s" % e
        
    # Performe cleanups
    try:
        gcsv.Close()
        mmdb.Close()
    except NameError, e:
        # Cleanup is performed according to creation order. When we encounter NameError,
        # we know for sure we cleaned all objects.
        pass
    except Exception, e:
        print "An error occurred after converting Google CSV file to Motorola mdb file!"
        print "Details: %s" % e

if __name__ == "__main__":
    main()