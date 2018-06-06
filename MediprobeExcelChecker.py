import sys
from TextShower import TextShower
from ExcelFormatter import ExcelFormatter

def PrintResults(errors_found):
	if errors_found >= 0:
		print(' ' + str(errors_found) + ' Errors Found.')

text_shower = 0
accepted_filename = False
output_filename = 'output.xlsx'
errors_found = -1

list_allowed_states = ['Gujarat', 'Madhya Pradesh', 'Maharashtra', 'Rajasthan']
list_allowed_cities = ['Achalpur', 'Ahmedabad', 'Ahmednagar', 'Ajmer', 'Akluj', 'Akola', 'Alibaug', 'Alwar', 'Amalner', 'Amaravati', 'Ambajogai', 'Ambarnath', 'Ambejogai', 'Amravati', 'Amreli', 'Anand', 'Andheri West', 'Anjangaon', 'Ankleshwar', 'Aravalli', 'Ardhapur', 'Aurangabad', 'Balghat', 'Banaskantha', 'Banswara', 'Baramati', 'Baran', 'Bardoli', 'Barmer', 'Baroda', 'Barshi', 'Bavla', 'Bayad', 'Beed', 'Bhandara', 'Bharatpur', 'Bharuch', 'Bhavnagar', 'Bhayander East', 'Bhilwara', 'Bhiwandi', 'Bhopal', 'Bhor', 'Bhuj', 'Bhuldhana', 'Bhusawal', 'Bikaner', 'Borivali West', 'Borsad', 'Buldana', 'Buldhana', 'Bundi', 'Chalisgaon', 'Chandrapur', 'Chembur', 'Chinchwad', 'Chiplun', 'Chittorgarh', 'Chota Udepur', 'Churu', 'Dadar', 'Dahigaon', 'Dahiwadi', 'Dahod', 'Daman', 'Damoh', 'Daryapur', 'Daund', 'Dausa', 'Deesa', 'Dewas', 'Dhar', 'Dharur', 'Dholka', 'Dhule', 'Dombivli', 'Dungarpur', 'Gadchiroli', 'Gandhidham', 'Gandhinagar', 'Godhra', 'Gondal', 'Gondawale', 'Gondia', 'Guna', 'Gwalior', 'Halol', 'Himmatnagar', 'Hinganghat', 'Hingoli', 'Hosangabad', 'Hupari', 'Ichalkaranji', 'Indapur', 'Indore', 'Islampur', 'Jabalpur', 'Jaipur', 'Jakhangaon', 'Jalgaon', 'Jalna', 'Jamnagar', 'Jath', 'Jaysingpur', 'Jejuri', 'Jhotwara', 'Jodhpur', 'Junagadh', 'Junnar', 'Kalol', 'Kalyan', 'Kalyan West', 'Kandivali East', 'Kandiwali', 'Karad', 'Karjan', 'Kasegaon', 'Katni', 'Khambhat', 'Khandala', 'Khandwa', 'Khargone', 'Khatav', 'Kheda', 'Khopoli', 'Kolhapur', 'Kopargaon', 'Kota', 'Kurduvadi', 'Kurla', 'Kusgaon Budruk', 'Kutch', 'Lashkar', 'Latur', 'Lonand', 'Lonand Satara', 'Lonavala', 'Lonavla', 'Majalgaon', 'Makronia', 'Malvan', 'Mandsaur', 'Mandvi', 'Mangaon', 'Mansa', 'Mehsana', 'Mhaswad', 'Mhow', 'Mira Bhayandar', 'Miraj', 'Modasa', 'Morbi', 'Mumbai', 'Murti', 'Nadiad', 'Nagaur', 'Nagpur', 'Nanded', 'Nandgaon', 'Nandurbar', 'Nanekarwadi', 'Nashik', 'Natepute', 'Navi Mumbai', 'Navsari', 'Neemuch', 'New Panvel ,Raigad', 'Nira', 'Osmanabad', 'PCMC', 'Padra', 'Palanpur', 'Palghar', 'Pali', 'Pandharpur', 'Panvel', 'Panvel West', 'Panvel Navi Mumbai', 'Parali', 'Parbhani', 'Parli', 'Patan', 'Perne', 'Pimpalgaon', 'Porbandar', 'Pune', 'Pusegaon', 'Raigad', 'Rajkot', 'Ratlam', 'Ratnagiri', 'Rewa', 'Sabarkantha', 'Sagar', 'Sangali', 'Sangli', 'Sangola', 'Saswad', 'Satara', 'Satna', 'Satpur Colony', 'Sawai Madhopur', 'Shahdol', 'Shindhudurg', 'Shirala', 'Shirur', 'Shirwal', 'Shrigonda', 'Shrirampur', 'Sikar', 'Sindhudurg', 'Sindkhed Raja', 'Sirsala', 'Solapur', 'Sri Ganganagar', 'Sumerpur', 'Surat', 'Surendra Nagar', 'Tapi', 'Taradgaon', 'Thane', 'Thane West', 'Tongaon', 'Tonk', 'Udaipur', 'Udgir', 'Ujjain', 'Ulhasnagar', 'Urun Islampur', 'Vadodara', 'Vaijapur', 'Valsad', 'Vapi', 'Varvand', 'Vasai', 'Veraval', 'Virar', 'Virar West', 'Vita', 'Walchandnagar', 'Wardha', 'Warwand', 'Wategaon', 'Yavatmal', 'Yawatmal', 'Karad', 'Kota', 'Saswad']

while not accepted_filename:
	text_shower = TextShower('Enter Filename (.xls/.xlsx)')
	text_shower.waitForInput()
	if text_shower.getString().endswith(('', 'xls', 'xlsx')):
		accepted_filename = True

if text_shower.getString() is not '':
	print('Starting Macro on "' + text_shower.getString() + '"...')
	excel_form = ExcelFormatter(text_shower.getString())
	print('\tChecking Address Lines...', end="", flush=True)
	errors_found = excel_form.CheckAddressLines(['Hospital_AddressLine1', 'Hospital_AddressLine2', 'Hospital_AddressLine3'])
	PrintResults(errors_found)

	print('\tChecking Email Addresses...', end="", flush=True)
	errors_found = excel_form.CheckEmailAddressLines(['Hospital_Email','Hospital_CEOEmail','Hospital_AdminEmail','Hospital_Insurance_Or_TPA_CoordinatorEmail','Hospital_Medical_SuperintendentEmail','Hospital_PromoterEmail'])
	PrintResults(errors_found)

	print('\tChecking Phone Numbers...', end="", flush=True)
	errors_found = excel_form.CheckPhoneNumbersLines(['Hospital_Phone','Hospital_CEOPhone','Hospital_AdminPhone','Hospital_Insurance_Or_TPA_CoordinatorPhone','Hospital_Medical_SuperintendentPhone','Hospital_PromoterPhone'])
	PrintResults(errors_found)

	print('\tChecking Fax Numbers...', end="", flush=True)
	errors_found = excel_form.CheckPhoneNumbersLines(['Hospital_Fax','Hospital_CEOFax','Hospital_PromoterFax'])
	PrintResults(errors_found)

	print('\tChecking Latitude Longitudes...', end="", flush=True)
	errors_found = excel_form.CheckLatLonLines(['Hospital_Latitude_Longitude'])
	PrintResults(errors_found)

	print('\tChecking Titles...', end="", flush=True)
	errors_found = excel_form.CheckTitleLines(['Hospital_CEOTitle','Hospital_AdminTitle','Hospital_Insurance_Or_TPA_CoordinatorTitle','Hospital_Medical_SuperintendentTitle','Hospital_PromoterTitle'])
	PrintResults(errors_found)

	print('\tChecking Hours...', end="", flush=True)
	errors_found = excel_form.CheckHoursLines(['Hospital_OPD_Working_Hours_StartTime','Hospital_OPD_Working_Hours_EndTime'])
	PrintResults(errors_found)

	print('\tChecking YY/NO/APP...', end="", flush=True)
	errors_found = excel_form.CheckYYNOAPPLines(['Hospital_HOTA_Registration','Hospital_PNDT_Registration','Hospital_Jci_Accredited','Hospital_Iso_Certified','Hospital_Nabh_Certified','Hospital_Lab_Nabl_Certified','Hospital_Other_Certification'])
	PrintResults(errors_found)

	print('\tChecking States Names...', end="", flush=True)
	errors_found = excel_form.CheckAllowedName(['Hospital_State'], list_allowed_states)
	PrintResults(errors_found)

	print('\tChecking Cities Names...', end="", flush=True)
	errors_found = excel_form.CheckAllowedName(['Hospital_City'], list_allowed_cities)
	PrintResults(errors_found)

	print('Saving output to ' + output_filename + '...', end="", flush=True)
	excel_form.SaveWorkbook(output_filename)
	print(' Completed!')
	print('Macro Completed!')


