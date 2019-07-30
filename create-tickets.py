# WebHelpDesk API Docs:
# http://www.solarwinds.com/documentation/WebHelpDesk/docs/WHD_API/Web%20Help%20Desk%20API%20Guide.html

import os
import sys
import requests
import xlrd
import json
from urllib.parse import quote
import keyring
import traceback

# If you used verify=False and receive an "Insecure Request" warning, suppress it by uncommenting the lines below
# import urllib3
# urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

site_url = "https://www.example.com/helpdesk/WebObjects/Helpdesk.woa/ra/" # URL for your Web Help Desk installation

# Uncomment one or both of the lines below to set default values
# ticket_file_path = "Bulk Ticket Creation Template.xlsx"
# tech_email_address = "email@example.com"

test_run = False # Change to True to print JSON instead of creating tickets

# Uncomment and populate with Location and ID pairs if you want to set defaults
# known_locations = {
# 	'Location 1': 1
# }

# Configure a few ticket defaults here
ticket_defaults = {
	"emailClient": True,
	"emailTech": True,
	"emailCc": False,
	"emailBcc": False,
	"assignToCreatingTech": False,
	"sendEmail": True,
}

# Colors and styles for emphasis on the command line
class colors:
	END = '\033[0m'
	BLUE = '\033[94m'
	GREEN = '\033[92m'
	YELLOW = '\033[33m'
	RED = '\033[31m'
	BOLD = '\033[1m'
	UNDERLINE = '\033[4m'

# Define the Custom Field IDs for the columns listed on your Excel spreadsheet
# If one Excel column corresponds to different custom fields depending on the ticket type, provide the ticket types and IDs as a dict
definition_ids = {
	"First Name": 1,
	"Last Name": 2,
	"Phone Number": 3,
	"Position": 4,
	"Yes/No?": 5,
	"Email for Client": 6
}

# Each of the ticket types listed below corresponds to a tab on the Excel spreadsheet.
#
# Define the following:
#	- The ticket type ID
#	- A list of Required fields using the names from definition_ids
#	- A list of Optional fields, also using the names from definition_ids. Leave empty if none.

supported_ticket_types = {
	"Example Ticket Type": {
		"id": 8, 
		"required": ["First Name", "Last Name", "Phone Number", "Position", "Yes/No?", "Email for Client"],
		"optional": []
	}
}

def query_api(url):
	response = requests.get(url, verify=True) # If the script is unable to verify the certificate, set verify=False
	response.raise_for_status()
	return(response)

def create_tickets(tickets):

	for index, ticket in enumerate(tickets):
		print("Submitting Ticket #%s - %s..." % (index+1, ticket['detail']))
	
		try:
			# Send the json data by POST
			response = requests.post(post_url, data=json.dumps(ticket), verify=True) # If the script is unable to verify the certificate, set verify=False
			response.raise_for_status()
	
			results = json.loads(response.text) #convert the response to a dict	
			print("%sCreated ticket #%s%s" % (colors.GREEN, results['id'], colors.END))
		except:
			print("%sUnable to create a ticket: %s%s" % (colors.RED, response.content.decode("utf-8"), colors.END))
	
def get_location(location):
	try:
		if "known_locations" in globals() and location in known_locations.keys():
			location_url = (site_url + "Locations/" + str(known_locations[location]) + "?apiKey=" + apiKey)
			response = query_api(location_url)
			
			results = json.loads(response.text)
			return(results)
		else:
			location_url = (site_url + "Locations?qualifier=" +
				"(locationName%20like%20'" + quote(location) + "')" +
				"&apiKey=%s" % apiKey)
			response = query_api(location_url)
					
			results = json.loads(response.text)

			location_count = len(results)

			if location_count == 0:
				raise SystemExit("Unable to find location: %s" % location)
			elif location_count == 1:
				return(results[0])
			elif location_count > 1:
				print("Multiple locations found matching '%s', please be more specific:\n" % location)
		
				locations = []
				for result in results:
					print("ID: %s, Location Name: %s" % (result['id'], result['locationName']))
					locations.append(result['id'])
		
				id = int(input("\nEnter the ID to use: "))
					
				if id in locations:
					location_url = (site_url + "Locations/" + str(id) + "?apiKey=" + apiKey)
					response = query_api(location_url)
					results = json.loads(response.text)
					return(results)
				else:
					raise SystemExit("Please enter a valid location ID.")
	except Exception as error:
		raise SystemExit("\n%sUnable to get location information: %s%s" % (colors.RED, error, colors.END))

def get_client(client):
	try:
		client_url = (site_url + "Clients?qualifier=" +
			"(email%20like%20'" + quote(client) + "*')" +
			"&apiKey=%s" % apiKey)	
		response = query_api(client_url)
		results = json.loads(response.text)
		
		client_count = len(results)

		if client_count == 0:
			return(False)
		elif client_count == 1:
			return(results[0])
		elif client_count > 1:
	
			print("Multiple clients found matching '%s', please be more specific:\n" % client)
	
			clients = []
			for result in results:
				print('ID: %s, Name: %s %s, Email: %s' % (result['id'], result['firstName'], result['lastName'], result['email']))
				clients.append(result['id'])
					
			id = int(input("\nEnter the ID to use: "))
		
			if id in clients:
				location_url = (site_url + "Clients/" + str(id) + "?apiKey=" + apiKey)
				response = query_api(location_url)
				results = json.loads(response.text)
				return(results)
			else:
				raise SystemExit("Please enter a valid client ID.")
	except Exception as error:
		raise SystemExit("\n%sUnable to get client information: %s%s" % (colors.RED, error, colors.END))

def construct_data(row):

	try:
		data = ticket_defaults
	
		data['customFields'] = []
		data['problemtype'] = {
			"id": supported_ticket_types[ticket_type]['id'], 
			"type": "RequestType"
		}
	
		required_fields = supported_ticket_types[ticket_type]['required']
	
		for field in required_fields:
		
			if (field == "Email for Client"): # Set this email address as the client user
				client_user = get_client(row[field])
				
				if not client_user:
					print("%sUnable to find client: %s%s" % (colors.RED, row[field], colors.END) )
					return(False)
				
				data['clientReporter'] = client_user # set the client user
				
				continue # skip the rest of this for loop
					
			if type(definition_ids[field]) is dict:
				try:
					data['customFields'].append({
						"definitionId": definition_ids[field][ticket_type],
						"restValue": row[field]
					})
				except KeyError:
					print("%s is missing information for %s in the Python script" % (field, ticket_type))
			else:
				data['customFields'].append({
					"definitionId": definition_ids[field],
					"restValue": row[field]
				})
	
		optional_fields = supported_ticket_types[ticket_type]['optional']
	
		if optional_fields:
			for field in optional_fields:
				if row[field]:
					data['customFields'].append({
						"definitionId": definition_ids[field],
						"restValue": row[field]
					})
		
		data['location'] = get_location(location)

		if 'Request Detail' in row.keys():
			data['detail'] = "[BULK TICKET] %s" % row['Request Detail']
		else:
			data['detail'] = "[BULK TICKET] %s for %s %s" % (ticket_type, row['First Name'], row['Last Name'])
	except Exception as error:
		print("\n%sUnable to construct data: %s%s" % (colors.RED, error, colors.END))
		
	return(data)
	
def get_rows(ticket_type):
	
	try:
		# read the header row as keys
		ticket_list = book.sheet_by_name(ticket_type)

		keys = ticket_list.row_values(0)

		row_count = ticket_list.nrows
	
		ticket_results = []
	
		if row_count > 1:

			print("\n%sFound %s ticket(s) for %s:%s" % (colors.BOLD, (row_count - 1), ticket_type, colors.END))

			for i in range(1, row_count):

				current_row = i + 1
					
				values = ticket_list.row_values(i)

				row = {}
				row['Ticket Type'] = ticket_type

				for key, value in zip(keys, values):
					row[key] = value

				required_fields = supported_ticket_types[ticket_type]['required']

				missing_fields = []

				for field in required_fields:
					try:
						if not row[field]:
							missing_fields.append(field)
					except KeyError:
						# The spreadsheet is missing a required column, so it must be an old version
						raise SystemExit("\n%sThe required column \"%s\" is missing. Please make sure you are using the most current version of the spreadsheet!%s\n" % (colors.RED, field, colors.END))
		
				if missing_fields:
					# Required fields were not populated on the spreadsheet
					print("%sRow #%s%s is missing the following required fields:\n%s%s%s" % (colors.BOLD, current_row, colors.END, colors.RED, ', '.join(missing_fields), colors.END))
					continue # skip the rest of this loop
				
				print("%sRow #%s%s: %s" % (colors.BOLD, current_row, colors.END, '\t'.join(map(str, row.values()))))
			
				data = construct_data(row) # gather all of the required information into a dict
			
				if data:
					ticket_results.append(data.copy())
			
			return ticket_results
											
	except Exception as error:
		# An unexpected error happened and we need to know more information
		exc_type, exc_obj, tb = sys.exc_info()
		lineno = tb.tb_lineno
		raise SystemExit("%sSorry, there was a problem on line %s: %s - %s%s" % (colors.RED, lineno, error, type(error), colors.END))
						
if __name__ == '__main__':

	if test_run:
		print("\n%sTHIS IS A TEST RUN. Submission will be simulated!%s\n" % (colors.BOLD, colors.END))

	if "tech_email_address" not in globals():
		tech_email_address = input("Email address for API user (tech account)? ").replace("\ ", " ").rstrip()
	
	try:
		apiKey = keyring.get_password("WebHelpDeskAPI", tech_email_address) # check for an existing API key
		
		if apiKey is None:
			apiKey = input("API key for %s? " % tech_email_address).replace("\ ", " ").rstrip()
			keyring.set_password("WebHelpDeskAPI", tech_email_address, apiKey) # add the API key to the system keychain
		
		post_url = "%sTickets?apiKey=%s" % (site_url, apiKey) # we have a key
		
	except Exception as error:
		raise SystemExit("\n%sI couldn't get the API key: %s%s" % (colors.RED, error, colors.END))

	if "ticket_file_path" not in globals():
		ticket_file_path = input("Path to ticket file? ").replace("\ ", " ").rstrip()
		
	filename, file_extension = os.path.splitext(ticket_file_path)

	if file_extension in ['.xls','.xlsx','.xlsm']:
		try:
			with xlrd.open_workbook(ticket_file_path) as book:
				print("\n%sOpening ticket file:%s %s\n" % (colors.BOLD, colors.END, ticket_file_path))
				
				location = book.sheet_by_name("Bulk Ticket Requests").cell(3, 4).value
				if not location:
					location = input("%sNo location specified!%s What location should I use? " % (colors.RED, colors.END)).rstrip()
					if not location:
						raise SystemExit("\nLocation is required!\n")
		
				client_name = book.sheet_by_name("Bulk Ticket Requests").cell(5, 4).value
				client_email = book.sheet_by_name("Bulk Ticket Requests").cell(7, 4).value
				if not client_email:
					client_email = input("%sNo client email specified!%s What email address should I use? " % (colors.RED, colors.END)).rstrip()
					if not client_email:
						raise SystemExit("\nClient email is required!\n")
			
				client_default = get_client(client_email)
			
				if not client_default:
					raise SystemExit("Unable to find client: %s" % client_email)
			
				ticket_defaults['clientReporter'] = client_default
				
				cc_email = book.sheet_by_name("Bulk Ticket Requests").cell(9, 4).value
				if cc_email:
					ticket_defaults['emailCc'] = True
					ticket_defaults['ccAddressesForTech'] = cc_email
			
				print("%sLocation:%s %s" % (colors.BOLD, colors.END, location))
				if client_name:
					print("%sClient name:%s %s" % (colors.BOLD, colors.END, client_name))
				print("%sClient email:%s %s" % (colors.BOLD, colors.END, client_email))
				
				all_tickets = []
								
				for ticket_type in supported_ticket_types.keys(): # look for rows in each sheet
					tickets = get_rows(ticket_type)
					if tickets:
						all_tickets += tickets
				
				ready = input("\n%sReady to submit %s tickets? [Y/n]%s" % (colors.BOLD, len(all_tickets), colors.END)).rstrip() or "y"
			
				if ready.lower() in ['y', 'yes']:
					if test_run:
						print(json.dumps(all_tickets)) # just print the data for test purposes
					else:
						create_tickets(all_tickets)
				else:
					print("Ticket submission cancelled for %s." % ticket_type)

									
		except Exception as error:
			# An unexpected error happened and we need to know more information
			exc_type, exc_obj, tb = sys.exc_info()
			lineno = tb.tb_lineno
			raise SystemExit("%sSorry, there was a problem on line %s: %s - %s%s" % (colors.RED, lineno, error, type(error), colors.END))
	else:
		raise SystemExit("Supported file types are .xls, .xlsx and .xlsm.")