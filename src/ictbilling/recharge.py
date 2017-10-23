from __future__ import print_function
import os
import copy

class GL:
	def __init__( self, company="IC", cost_centre="ITSO", activity="G80447", account="163116" ):
		if not isinstance (company, str) :
			raise Exception("company is not a string")
		if not isinstance( cost_centre, str) :
			raise Exception("cost_centre is not a string")
		if not isinstance( activity, str) :
			raise Exception("activity is not a string")
		if not isinstance(account, str) :
			raise Exception("account is not a string")
		self.company     = company
		self.cost_centre = cost_centre
		self.activity    = activity
		self.account     = account
	
	def __str__(self):
		return "%s-%s-%s-%s" % ( self.company, self.cost_centre, self.activity, self.account )

class Recharge:
	def __init__(self, description=None, destination=None):
		if not isinstance( description, str ) :
			raise Exception("description is not a string")
		if not isinstance( destination, GL ):
			raise Exception("destination is not a GL object")
		self.description = description
		self.destination = destination
		self.transactions= []

	def addTransaction( self, source=None, amount=None, description=None ):
		if not isinstance( source, GL ):
			raise Exception("source is not a GL object")
		if not isinstance( description, str ) :
			raise Exception("description is not a string")

		amount = float( amount )
		if amount <= 0.0:
			raise Exception( "amount is invalid" )

		txn = copy.deepcopy(source)
		txn.amount      = amount
		txn.description = description
		self.transactions.append( txn )

	def writeOutput( self, filename ):
		import openpyxl
		import inspect
		# Load up the template
		f = inspect.getfile( GL )
		f = os.path.dirname( f )
		f = os.path.join( f, "data", "WEB-ADI-template.xlsm" )
		wb = openpyxl.load_workbook( filename = f, read_only = False, keep_vba = True )
		sheet = wb["Sheet1"]


		# copy footer
		for row in [25, 26]: 
			for col in "ABCDEFGHIJKLMNO":
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].value = sheet[ "%s%d" % ( col, row ) ].value
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].style = sheet[ "%s%d" % ( col, row) ].style
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].border    = copy.copy( sheet[ "%s%d" % ( col, row) ].border )
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].number_format    = copy.copy( sheet[ "%s%d" % ( col, row) ].number_format )
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].fill      = copy.copy( sheet[ "%s%d" % ( col, row) ].fill )
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].alignment = copy.copy( sheet[ "%s%d" % ( col, row) ].alignment )
				sheet["%s%d" % ( col, row + len(self.transactions)*2-2) ].font  = copy.copy( sheet[ "%s%d" % ( col, row) ].font )


		sheet[ "J%d" % (23 + len(self.transactions)*2) ].value = "=SUM(J23:J%d)" % ( 24 + len(self.transactions)*2 - 2)
		sheet[ "K%d" % (23 + len(self.transactions)*2) ].value = "=SUM(K23:K%d)" % ( 24 + len(self.transactions)*2 - 2)

		for row in range( 1, len(self.transactions )):
			for i in [23, 24]:
				for col in "ABCDEFGHIJKLMNO":
					sheet["%s%d" % ( col, i + (row * 2 )) ].value = sheet[ "%s%d" % ( col, i  ) ].value
					sheet["%s%d" % ( col, i + (row * 2 )) ].style = sheet[ "%s%d" % ( col, i  ) ].style
					sheet["%s%d" % ( col, i + (row * 2 )) ].border = copy.copy( sheet[ "%s%d" % ( col, i ) ].border )
					sheet["%s%d" % ( col, i + (row * 2 )) ].number_format = copy.copy( sheet[ "%s%d" % ( col, i ) ].number_format )
					sheet["%s%d" % ( col, i + (row * 2 )) ].fill   = copy.copy( sheet[ "%s%d" % ( col, i ) ].fill )
					sheet["%s%d" % ( col, i + (row * 2 )) ].alignment = copy.copy( sheet[ "%s%d" % ( col, i ) ].alignment )
					sheet["%s%d" % ( col, i + (row * 2 )) ].font  = copy.copy( sheet[ "%s%d" % ( col, i ) ].font )
	
		# copy in transactions
		
		idx = 23
		
		for t in self.transactions:
#			sheet[ "B%d" % ( idx ) ] = "C" 
			sheet[ "C%d" % ( idx ) ] = t.company
			sheet[ "D%d" % ( idx ) ] = t.cost_centre
			sheet[ "E%d" % ( idx ) ] = t.activity
			sheet[ "F%d" % ( idx ) ] = t.account
			sheet[ "G%d" % ( idx ) ] = "0"
			sheet[ "H%d" % ( idx ) ] = "0"
			sheet[ "I%d" % ( idx ) ] = "0"
			sheet[ "J%d" % ( idx ) ] = t.amount
			sheet[ "K%d" % ( idx ) ] = ""
			sheet[ "L%d" % ( idx ) ] = t.description
			idx = idx + 1

#			sheet[ "B%d" % ( idx ) ] = "C" 
			sheet[ "C%d" % ( idx ) ] = t.company
			sheet[ "D%d" % ( idx ) ] = t.cost_centre
			sheet[ "E%d" % ( idx ) ] = t.activity
			sheet[ "F%d" % ( idx ) ] = t.account
			sheet[ "G%d" % ( idx ) ] = "R"
			sheet[ "H%d" % ( idx ) ] = "0"
			sheet[ "I%d" % ( idx ) ] = "0"
			sheet[ "J%d" % ( idx ) ] = ""
			sheet[ "K%d" % ( idx ) ] = t.amount
			sheet[ "L%d" % ( idx ) ] = "Recharge for " + t.description + "(" + str(t) + ")"
			idx = idx + 1

		wb.save( filename = filename )



if __name__ == "__main__":
	print("Hello")	
	r = Recharge( description="X" , destination = GL() )
	src = GL( cost_centre="EDTA", activity = "G80000", account = "123456" )	
	r.addTransaction( src, 100.0, "Recharge amount" )

	src = GL( cost_centre="XDTA", activity = "G81235", account = "987654" )	
	r.addTransaction( src, 500.0, "Recharge amount for compute" )
	r.writeOutput( "/tmp/xx.xlsm" )
