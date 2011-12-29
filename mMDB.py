import win32com.client
import os

NON_ENG_LANGCODE = 3 # Hebrew

def is_ascii(s):
    return all(ord(c) < 128 for c in s)

"""
This class implements motorola phonebook manager database.
You may add contacts using this class.
"""
class mMDB:
    def __init__(self,mdb_file_path,override_mdb_file=False):
        # Set mmdb file path as class property 
        self.mdb_file_path = mdb_file_path
        
        # Create mmdb engine instance
        db_engine = win32com.client.gencache.EnsureDispatch("DAO.DBEngine.35")
        
        # Check whether mmdb file already exists
        if os.path.exists(self.mdb_file_path):
            # If asked to override, delete file and create a new mmdb
            if override_mdb_file:
                os.remove(self.mdb_file_path)
                db = db_engine.CreateDatabase(mdb_file_path, win32com.client.constants.dbLangGeneral)
                self._InitMmdb(db,True)
            # Just open existing mmdb file
            else:
                db = db_engine.OpenDatabase(mdb_file_path)
                self._InitMmdb(db,False)
        # mmdb doesn't exists. Create a new one.
        else:
                db = db_engine.CreateDatabase(mdb_file_path, win32com.client.constants.dbLangGeneral)
                self._InitMmdb(db, True)
        
    def _InitMmdb(self,mmdb,is_new_mmdb):
        self.mmdb = mmdb
        if is_new_mmdb:
            # Create DB Tables
            self.mmdb.Execute("CREATE TABLE tblContactList (ID Integer, ContactID Integer, ContactName Text (22), SpeedDial Integer, [Number] Text (201), NumberType Integer, Language Byte, NameLang Byte, PhoneNumberLength Byte, Reserved Integer, TonNP Byte, CapabilityID Byte, ExtnRecID Byte, VoiceTag Integer, TStamp0 Byte, TStamp1 Byte, TStamp2 Byte, TStamp3 Byte, SyncStatus Integer, TempID Integer, RingerID Byte, PictureID Integer, ExtnRecIDMSB Byte)")
            self.mmdb.Execute("CREATE TABLE tblGSMPhonebook ([Index] Integer, AliasName Text (12), PhoneNumber Text (21), Programmed Byte, Language Byte)")
            self.mmdb.Execute("CREATE TABLE tblIdentification (ID Integer, Title Text (50), SIM_Model Text (50), LastUpDate DateTime, TotalContacts Integer, TotalEntries Integer, ListSize Integer, Recognition Text (24))")
            self.mmdb.Execute("CREATE TABLE tblNumberTypes (ID Integer, TypeName Text (50), [Value] Integer)")
            
            # Insert Identification defaults
            self.mmdb.Execute("INSERT INTO tblIdentification VALUES (1,\"IPM\",\"\",\"09/24/03 00:00:00\",0,0,0,\"iDEN Phonebook Manager\")")
        
            # Insert Number types data
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (1,\"Main\",0)")       
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (2,\"Private ID\",1)")       
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (3,\"Home\",2)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (4,\"Work 1\",3)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (5,\"Mobile\",4)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (6,\"Fax\",5)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (7,\"Pager\",6)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (8,\"Talkgroup\",7)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (9,\"Chat Address\",8)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (10,\"Other\",9)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (11,\"Email\",10)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (12,\"Work 2\",11)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (13,\"Private ID2\",13)")
            self.mmdb.Execute("INSERT INTO tblNumberTypes VALUES (14,\"Email2\",14)")
            
            # Initialize tblContactList counters
            self.ContactId = 1
            self.RecordId = 1
        else:
            # Get latest Record ID and Contact ID
            recordset = self.mmdb.OpenRecordset('SELECT MAX([ID]), MAX(ContactID) FROM tblContactList')
            res = recordset.GetRows()
            
            self.RecordId = res[0][0]
            self.ContactId = res[1][0]
            
            # Increase Counters by 1 for correct count
            # If no records present, set both to 1
            if self.RecordId:
                self.RecordId = self.RecordId + 1
                self.ContactId = self.ContactId + 1
            else:
                self.RecordId = 1
                self.ContactId = 1
            
    def _AddContactRecord(self,ContactName,PhoneText,PhoneType):
            if is_ascii(ContactName):
                LanguageCode = 0
            else:     
                LanguageCode = NON_ENG_LANGCODE
            
            self.mmdb.Execute("INSERT INTO tblContactList VALUES (%s,%s,\"%s\",%s,\"%s\",%s,%s,%s,%s,0,0,144,0,0,0,0,0,0,0,0,255,0,0)" % (self.RecordId,self.ContactId, ContactName, self.RecordId, PhoneText, PhoneType, LanguageCode, LanguageCode, len(PhoneText)))
            self.RecordId = self.RecordId + 1
            
    def AddContact(self,contact):        
        # Now we add contact details according to the following scheme:
        # ID,ContactID,ContactName,SpeedDial,Number,NumberType,Language,NameLang,PhoneNumberLength,Reserved,TonNP,CapabilityID,ExtnRecID,VoiceTag,TStamp0,TStamp1,TStamp2,TStamp3,SyncStatus,TempID,RingerID,PictureID,ExtnRecIDMSB
        if contact.Home: self._AddContactRecord(contact.Name,contact.Home,2)
        if contact.Mobile: self._AddContactRecord(contact.Name,contact.Mobile,4)
        if contact.Work1: self._AddContactRecord(contact.Name,contact.Work1,3)
        if contact.Work2: self._AddContactRecord(contact.Name,contact.Work2,11)
        if contact.Email: self._AddContactRecord(contact.Name,contact.Email,10)
        if contact.Email2: self._AddContactRecord(contact.Name,contact.Email2,14)
        if contact.Fax: self._AddContactRecord(contact.Name,contact.Fax,5)
        
        self.ContactId = self.ContactId + 1
            
    def AddContacts(self,contacts):
        for contact in contacts: self.AddContact(contact)
                 
    def Close(self):
        self.mmdb.Close()