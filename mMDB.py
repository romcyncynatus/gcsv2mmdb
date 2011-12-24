import win32com.client
import os

"""
This class implements motorola phonebook manager database.
You may add contacts using this class.
"""

class mMDB:    
    def __init__(self,mdb_file_path):
        if os.path.exists(mdb_file_path):
            raise Exception, "mdb output file already exists."
    
        db_eng = win32com.client.gencache.EnsureDispatch("DAO.DBEngine.35")
        self.db = db_eng.CreateDatabase(mdb_file_path, win32com.client.constants.dbLangGeneral)
        
        # Create DB Tables
        self.db.Execute("CREATE TABLE tblContactList (ID Integer, ContactID Integer, ContactName Text (22), SpeedDial Integer, [Number] Text (201), NumberType Integer, Language Byte, NameLang Byte, PhoneNumberLength Byte, Reserved Integer, TonNP Byte, CapabilityID Byte, ExtnRecID Byte, VoiceTag Integer, TStamp0 Byte, TStamp1 Byte, TStamp2 Byte, TStamp3 Byte, SyncStatus Integer, TempID Integer, RingerID Byte, PictureID Integer, ExtnRecIDMSB Byte)")
        
        self.db.Execute("CREATE TABLE tblGSMPhonebook ([Index] Integer, AliasName Text (12), PhoneNumber Text (21), Programmed Byte, Language Byte)")

        self.db.Execute("CREATE TABLE tblIdentification (ID Integer, Title Text (50), SIM_Model Text (50), LastUpDate DateTime, TotalContacts Integer, TotalEntries Integer, ListSize Integer, Recognition Text (24))")

        self.db.Execute("CREATE TABLE tblNumberTypes (ID Integer, TypeName Text (50), [Value] Integer)")
        
        # Insert Identification defaults
        self.db.Execute("INSERT INTO tblIdentification VALUES (1,\"IPM\",\"\",\"09/24/03 00:00:00\",0,0,0,\"iDEN Phonebook Manager\")")
    
        # Insert Number types data
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (1,\"Main\",0)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (2,\"Private ID\",1)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (3,\"Home\",2)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (4,\"Work 1\",3)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (5,\"Mobile\",4)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (6,\"Fax\",5)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (7,\"Pager\",6)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (8,\"Talkgroup\",7)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (9,\"Chat Address\",8)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (10,\"Other\",9)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (11,\"Email\",10)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (12,\"Work 2\",11)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (13,\"Private ID2\",13)")
        
        self.db.Execute("INSERT INTO tblNumberTypes VALUES (14,\"Email2\",14)")
        
        # Initialize tblContactList counters
        self.ContactId = 1
        self.RecordId = 1
        
    def AddContact(self,ContactName,Language="English",Home=None,Work1=None,Work2=None,Mobile=None,Email=None,Fax=None):
        # We want to make sure that we wont increment the contacts counter
        # by mistake (if someone called us with a name only...)
        ContactAdded = False
        
        # Set language number according to input 
        LanguageCode = 0
        if Language == "Hebrew":
            LanguageCode = 3
        
        # Now we add contact details according to the following scheme:
        # ID,ContactID,ContactName,SpeedDial,Number,NumberType,Language,NameLang,PhoneNumberLength,Reserved,TonNP,CapabilityID,ExtnRecID,VoiceTag,TStamp0,TStamp1,TStamp2,TStamp3,SyncStatus,TempID,RingerID,PictureID,ExtnRecIDMSB
        
        if Home:
            Home = filter(lambda x: x.isdigit(), Home)
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",2,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Home, LanguageCode, LanguageCode, len(Home)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if Work1:
            Work1 = filter(lambda x: x.isdigit(), Work1)
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",3,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Work1, LanguageCode, LanguageCode, len(Work1)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if Work2:
            Work2 = filter(lambda x: x.isdigit(), Work2)
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",11,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Work2, LanguageCode, LanguageCode, len(Work2)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if Mobile:
            Mobile = filter(lambda x: x.isdigit(), Mobile)
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",4,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Mobile, LanguageCode, LanguageCode, len(Mobile)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if Email:
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",10,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Email, LanguageCode, LanguageCode, len(Email)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if Fax:
            Fax = filter(lambda x: x.isdigit(), Fax)
            self.db.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",5,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, Fax, LanguageCode, LanguageCode, len(Fax)))
            ContactAdded = True
            self.RecordId = self.RecordId + 1
            
        if ContactAdded:
            # As explained earlier, if we indeed added the contact, we increment the counter.
            self.ContactId = self.ContactId + 1
                 
    def Close(self):
        self.db.Close()