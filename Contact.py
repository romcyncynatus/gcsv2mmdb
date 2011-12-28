"""
This class represents a single contact
"""
class EmptyContactError(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)

class Contact:
	def __init__(self,Name,Home=None,Mobile=None,Work1=None,Work2=None,Email=None,Email2=None,Fax=None):
		self.Name = Name
		
		if Home:
			self.Home = filter(lambda x: x.isdigit(), Home) 
		else:
			self.Home = Home
			
		if Mobile:	
			self.Mobile = filter(lambda x: x.isdigit(), Mobile)
		else:
			self.Mobile = Mobile
			
		if Work1:
			self.Work1 = filter(lambda x: x.isdigit(), Work1)
		else:
			self.Work1 = Work1
		
		if Work2:
			self.Work2 = filter(lambda x: x.isdigit(), Work2)
		else:
			self.Work2 = Work2
			
		self.Email = Email
		self.Email2 = Email2
		
		if Fax:
			self.Fax = filter(lambda x: x.isdigit(), Fax)
		else:
			self.Fax = Fax
		
		# We check that at least one property was added
		has_properties = False	
		for property in (self.Home,self.Mobile,self.Work1,self.Work2,self.Email,self.Email2,self.Fax):
			if property: has_properties = True
		
		# Error if not	
		if not has_properties:
			raise EmptyContactError("Contact %s was being created without properties." % self.Name)			
		
