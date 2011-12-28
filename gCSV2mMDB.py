from gCSV import gCSV
from mMDB import mMDB
import sys

ARGS_COUNT = 3

def print_usage():
    print "Usage: %s <input csv file> <output mdb file>" % sys.argv[0]

def main():
    # Validate program arguments. Print usage and terminate upon incorrect usage.
    if len(sys.argv) != ARGS_COUNT:
        print_usage()
        sys.exit(1)
    
    # Create data objects and convert gcsv to mmdb
    try:
        # Create google csv file instance using the input file.
        gcsv = gCSV(sys.argv[1])
        
        # Create motorola mdb file instance for the output file. Override if exists.
        mmdb = mMDB(sys.argv[2],True)
        
        # Convert gcsv to mmdb
        contacts = gcsv.GetContacts()
        mmdb.AddContacts(contacts)
        
        # All went fine, let the world know :)
        print "Motorola mdb file was created successfully."
        print "Import your contacts using Motorola PhoneBook Manager."
    
    except Exception, e:
        print "An error occurred while converting Google CSV file to Motorola mdb file!"
        print "Details: %s" % e
        
    # Perform cleanups
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