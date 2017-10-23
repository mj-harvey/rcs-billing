class GL:
	def __init__( cost_centre=None, cost_code=None, activity=None, account=None ):
	


class Recharge:
	def __init__(self, description=None, destination=None):
		if description is not  instanceof str :
			raise Exception("description is not a string")
		if destination is not instanceof GL:
			raise Exception("destination is not a GL object")
		self.description = description
		self.destination = destination
		self.transactions= []

	def addTransaction( self, source=None, amount=None, description=None ):
		if source is not instanceof GL:
			raise Exception("source is not a GL object")
		if description is not  instanceof str :
			raise Exception("description is not a string")

		amount = (float) amount
		if amount <= 0.0:
			raise Exception( "amount is invalid" )

		txn = source.copy()
		txn.amount = amount
		self.transactions.append( txn )

	def writeOutput( self, filename ):
		import openpyxl
		# Load up the template
		
