import sys
from TextShower import TextShower
from ExcelFormatter import ExcelFormatter

def PrintResults(errors_found):
	if errors_found > 0:
		print(' ' + str(errors_found) + ' Errors Found.')
	else:
		print(" Completed!")

text_shower = 0
accepted_filename = False
output_filename = 'output.xlsx'

while not accepted_filename:
	text_shower = TextShower('Enter Filename (.xls/.xlsx)')
	text_shower.waitForInput()
	if text_shower.getString().endswith(('', 'xls', 'xlsx')):
		accepted_filename = True

if text_shower.getString() is not '':
	print('Starting Macro on "' + text_shower.getString() + '"...')
	excel_form = ExcelFormatter(text_shower.getString())
	print('Checking Address Lines...', end="", flush=True)
	errors_found = excel_form.CheckAddressLines(['Hospital_AddressLine1', 'Hospital_AddressLine2', 'Hospital_AddressLine3'])
	PrintResults(errors_found)

	print('Checking Email Addresses...', end="", flush=True)
	errors_found = excel_form.CheckEmailAddressLines(['Hospital_Email','Hospital_CEOEmail','Hospital_AdminEmail','Hospital_Insurance_Or_TPA_CoordinatorEmail','Hospital_Medical_SuperintendentEmail','Hospital_PromoterEmail'])
	PrintResults(errors_found)

	print('Checking Phone Numbers...', end="", flush=True)
	errors_found = excel_form.CheckPhoneNumbersLines(['Hospital_Phone','Hospital_CEOPhone','Hospital_AdminPhone','Hospital_Insurance_Or_TPA_CoordinatorPhone','Hospital_Medical_SuperintendentPhone','Hospital_PromoterPhone'])
	PrintResults(errors_found)

	print('Checking Fax Numbers...', end="", flush=True)
	errors_found = excel_form.CheckPhoneNumbersLines(['Hospital_Fax','Hospital_CEOFax','Hospital_PromoterFax'])
	PrintResults(errors_found)

	print('Checking Latitude Longitudes...', end="", flush=True)
	errors_found = excel_form.CheckLatLonLines(['Hospital_Latitude_Longitude'])
	PrintResults(errors_found)

	print('Saving output to ' + output_filename + '...', end="", flush=True)
	excel_form.SaveWorkbook(output_filename)
	print(' Completed!')
	print('Macro Completed!')


