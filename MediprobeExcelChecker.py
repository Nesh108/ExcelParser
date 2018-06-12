import sys
from ExcelFormatter import ExcelFormatter
from itertools import islice
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

def PrintResults(errors_found):
	if errors_found >= 0:
		print(' ' + str(errors_found) + ' Errors Found.')

input_filename = ""
output_filename = ""
errors_found = -1

list_allowed_states = ['Gujarat', 'Madhya Pradesh', 'Maharashtra', 'Rajasthan']
list_allowed_cities = ['Achalpur', 'Ahmedabad', 'Ahmednagar', 'Ajmer', 'Akluj', 'Akola', 'Alibaug', 'Alwar', 'Amalner', 'Amaravati', 'Ambajogai', 'Ambarnath', 'Ambejogai', 'Amravati', 'Amreli', 'Anand', 'Andheri West', 'Anjangaon', 'Ankleshwar', 'Aravalli', 'Ardhapur', 'Aurangabad', 'Balghat', 'Banaskantha', 'Banswara', 'Baramati', 'Baran', 'Bardoli', 'Barmer', 'Baroda', 'Barshi', 'Bavla', 'Bayad', 'Beed', 'Bhandara', 'Bharatpur', 'Bharuch', 'Bhavnagar', 'Bhayander East', 'Bhilwara', 'Bhiwandi', 'Bhopal', 'Bhor', 'Bhuj', 'Bhuldhana', 'Bhusawal', 'Bikaner', 'Borivali West', 'Borsad', 'Buldana', 'Buldhana', 'Bundi', 'Chalisgaon', 'Chandrapur', 'Chembur', 'Chinchwad', 'Chiplun', 'Chittorgarh', 'Chota Udepur', 'Churu', 'Dadar', 'Dahigaon', 'Dahiwadi', 'Dahod', 'Daman', 'Damoh', 'Daryapur', 'Daund', 'Dausa', 'Deesa', 'Dewas', 'Dhar', 'Dharur', 'Dholka', 'Dhule', 'Dombivli', 'Dungarpur', 'Gadchiroli', 'Gandhidham', 'Gandhinagar', 'Godhra', 'Gondal', 'Gondawale', 'Gondia', 'Guna', 'Gwalior', 'Halol', 'Himmatnagar', 'Hinganghat', 'Hingoli', 'Hosangabad', 'Hupari', 'Ichalkaranji', 'Indapur', 'Indore', 'Islampur', 'Jabalpur', 'Jaipur', 'Jakhangaon', 'Jalgaon', 'Jalna', 'Jamnagar', 'Jath', 'Jaysingpur', 'Jejuri', 'Jhotwara', 'Jodhpur', 'Junagadh', 'Junnar', 'Kalol', 'Kalyan', 'Kalyan West', 'Kandivali East', 'Kandiwali', 'Karad', 'Karjan', 'Kasegaon', 'Katni', 'Khambhat', 'Khandala', 'Khandwa', 'Khargone', 'Khatav', 'Kheda', 'Khopoli', 'Kolhapur', 'Kopargaon', 'Kota', 'Kurduvadi', 'Kurla', 'Kusgaon Budruk', 'Kutch', 'Lashkar', 'Latur', 'Lonand', 'Lonand Satara', 'Lonavala', 'Lonavla', 'Majalgaon', 'Makronia', 'Malvan', 'Mandsaur', 'Mandvi', 'Mangaon', 'Mansa', 'Mehsana', 'Mhaswad', 'Mhow', 'Mira Bhayandar', 'Miraj', 'Modasa', 'Morbi', 'Mumbai', 'Murti', 'Nadiad', 'Nagaur', 'Nagpur', 'Nanded', 'Nandgaon', 'Nandurbar', 'Nanekarwadi', 'Nashik', 'Natepute', 'Navi Mumbai', 'Navsari', 'Neemuch', 'New Panvel ,Raigad', 'Nira', 'Osmanabad', 'PCMC', 'Padra', 'Palanpur', 'Palghar', 'Pali', 'Pandharpur', 'Panvel', 'Panvel West', 'Panvel Navi Mumbai', 'Parali', 'Parbhani', 'Parli', 'Patan', 'Perne', 'Pimpalgaon', 'Porbandar', 'Pune', 'Pusegaon', 'Raigad', 'Rajkot', 'Ratlam', 'Ratnagiri', 'Rewa', 'Sabarkantha', 'Sagar', 'Sangali', 'Sangli', 'Sangola', 'Saswad', 'Satara', 'Satna', 'Satpur Colony', 'Sawai Madhopur', 'Shahdol', 'Shindhudurg', 'Shirala', 'Shirur', 'Shirwal', 'Shrigonda', 'Shrirampur', 'Sikar', 'Sindhudurg', 'Sindkhed Raja', 'Sirsala', 'Solapur', 'Sri Ganganagar', 'Sumerpur', 'Surat', 'Surendra Nagar', 'Tapi', 'Taradgaon', 'Thane', 'Thane West', 'Tongaon', 'Tonk', 'Udaipur', 'Udgir', 'Ujjain', 'Ulhasnagar', 'Urun Islampur', 'Vadodara', 'Vaijapur', 'Valsad', 'Vapi', 'Varvand', 'Vasai', 'Veraval', 'Virar', 'Virar West', 'Vita', 'Walchandnagar', 'Wardha', 'Warwand', 'Wategaon', 'Yavatmal', 'Yawatmal', 'Karad', 'Kota', 'Saswad']
list_numbers_ranges = [['Hospital_Beds_General', 'Hospital_Beds_General_Range'],['Hospital_Beds_Twin_Share', 'Hospital_Beds_Twin_Share_Range'],['Hospital_Beds_Single', 'Hospital_Beds_Single_Range'],['Hospital_Beds_ICU_NICU', 'Hospital_Beds_ICU_NICU_Range'],['Hospital_Beds_Single_AC', 'Hospital_Beds_Single_AC_Range'],['Hospital_Beds_Twin_Share_AC', 'Hospital_Beds_Twin_Share_AC_Range'],['Hospital_Beds_Suite', 'Hospital_Beds_Suite_Range'],['Hospital_Beds_Delux', 'Hospital_Beds_Delux_Range'],['Hospital_Beds_Others', 'Hospital_Beds_Others_Range'],['Hospital_Beds_Day_Care', 'Hospital_Beds_Day_Care_Range'],['Hospital_Beds_Total', 'Hospital_Beds_Total_Range'],['Hospital_NursesGeneral_Ward', 'Hospital_NursesGeneral_Ward_Range'],['Hospital_NursesTwin_Share', 'Hospital_NursesTwin_Share_Range'],['Hospital_NursesSingle', 'Hospital_NursesSingle_Range'],['Hospital_NursesICU', 'Hospital_NursesICU_Range'],['Hospital_NursesSingle_AC', 'Hospital_NursesSingle_AC_Range'],['Hospital_NursesTwin_Share_AC', 'Hospital_NursesTwin_Share_AC_Range'],['Hospital_NursesSuite', 'Hospital_NursesSuite_Range'],['Hospital_Nurses_Delux', 'Hospital_Nurses_Delux_Range'],['Hospital_Nurses_Others', 'Hospital_Nurses_Others_Range'],['Hospital_NursesDay_Care', 'Hospital_NursesDay_Care_Range'],['Hospital_NursesTotal', 'Hospital_NursesTotal_Range'],['Hospital_Resident_Doctors', 'Hospital_Resident_Doctors_Range'],['Hospital_Resident_Specialists', 'Hospital_Resident_Specialists_Range'],['Hospital_Visiting_Consultants', 'Hospital_Visiting_Consultants_Range'],['Hospital_PhysicianSpeciality_Internal_Medicine', 'Hospital_PhysicianSpeciality_Internal_Medicine_Range'],['Hospital_PhysicianSpeciality_Cardiology', 'Hospital_PhysicianSpeciality_Cardiology_Range'],['Hospital_PhysicianSpeciality_Nephrology', 'Hospital_PhysicianSpeciality_Nephrology_Range'],['Hospital_PhysicianSpeciality_Neonatology', 'Hospital_PhysicianSpeciality_Neonatology_Range'],['Hospital_PhysicianSpeciality_Pediatrics', 'Hospital_PhysicianSpeciality_Pediatrics_Range'],['Hospital_PhysicianSpeciality_Pulmonology', 'Hospital_PhysicianSpeciality_Pulmonology_Range'],['Hospital_PhysicianSpeciality_Gastro_Entomology', 'Hospital_PhysicianSpeciality_Gastro_Entomology_Range'],['Hospital_PhysicianSpeciality_General_Surgery', 'Hospital_PhysicianSpeciality_General_Surgery_Range'],['Hospital_PhysicianSpeciality_Orthopedics', 'Hospital_PhysicianSpeciality_Orthopedics_Range'],['Hospital_PhysicianSpeciality_Gynecology', 'Hospital_PhysicianSpeciality_Gynecology_Range'],['Hospital_PhysicianSpeciality_Obstetrics', 'Hospital_PhysicianSpeciality_Obstetrics_Range'],['Hospital_PhysicianSpeciality_Medical_Oncology', 'Hospital_PhysicianSpeciality_Medical_Oncology_Range'],['Hospital_PhysicianSpeciality_Surgical_Oncology', 'Hospital_PhysicianSpeciality_Surgical_Oncology_Range'],['Hospital_PhysicianSpeciality_Radiation_Oncology', 'Hospital_PhysicianSpeciality_Radiation_Oncology_Range'],['Hospital_PhysicianSpeciality_Urology', 'Hospital_PhysicianSpeciality_Urology_Range'],['Hospital_PhysicianSpeciality_Cardiothoracic', 'Hospital_PhysicianSpeciality_Cardiothoracic_Range'],['Hospital_PhysicianSpeciality_Ent', 'Hospital_PhysicianSpeciality_Ent_Range'],['Hospital_PhysicianSpeciality_Endocrinology', 'Hospital_PhysicianSpeciality_Endocrinology_Range'],['Hospital_PhysicianSpeciality_Ophthalmology', 'Hospital_PhysicianSpeciality_Ophthalmology_Range'],['Hospital_PhysicianSpeciality_Haematology', 'Hospital_PhysicianSpeciality_Haematology_Range'],['Hospital_PhysicianSpeciality_Rheumatology', 'Hospital_PhysicianSpeciality_Rheumatology_Range'],['Hospital_PhysicianSpeciality_Neurology', 'Hospital_PhysicianSpeciality_Neurology_Range'],['Hospital_PhysicianSpeciality_Neuro_Surgery', 'Hospital_PhysicianSpeciality_Neuro_Surgery_Range'],['Hospital_PhysicianSpeciality_Plastic_Surgery', 'Hospital_PhysicianSpeciality_Plastic_Surgery_Range'],['Hospital_PhysicianSpeciality_Vascular_Surgery', 'Hospital_PhysicianSpeciality_Vascular_Surgery_Range'],['Hospital_Number_Consulting_Room', 'Hospital_Number_Consulting_Room_Range'],['Hospital_Number_Major_Operating_Theaters', 'Hospital_Number_Major_Operating_Theaters_Range'],['Hospital_Number_Minor_Operating_Theaters', 'Hospital_Number_Minor_Operating_Theaters_Range']]
list_specialities = [['Hospital_Specialities_Internal_Medicine', 'Hospital_PhysicianSpeciality_Internal_Medicine', 'Hospital_PhysicianSpeciality_Internal_Medicine_Range'], ['Hospital_Specialities_Cardiology', 'Hospital_PhysicianSpeciality_Cardiology', 'Hospital_PhysicianSpeciality_Cardiology_Range'], ['Hospital_Specialities_Nephrology', 'Hospital_PhysicianSpeciality_Nephrology', 'Hospital_PhysicianSpeciality_Nephrology_Range'], ['Hospital_Specialities_Neonatology', 'Hospital_PhysicianSpeciality_Neonatology', 'Hospital_PhysicianSpeciality_Neonatology_Range'], ['Hospital_Specialities_Pediatrics', 'Hospital_PhysicianSpeciality_Pediatrics', 'Hospital_PhysicianSpeciality_Pediatrics_Range'], ['Hospital_Specialities_PulmoNologists', 'Hospital_PhysicianSpeciality_Pulmonology', 'Hospital_PhysicianSpeciality_Pulmonology_Range'], ['Hospital_Specialities_Gastro_Enterology', 'Hospital_PhysicianSpeciality_Gastro_Entomology', 'Hospital_PhysicianSpeciality_Gastro_Entomology_Range'], ['Hospital_Specialities_General_Surgery', 'Hospital_PhysicianSpeciality_General_Surgery', 'Hospital_PhysicianSpeciality_General_Surgery_Range'], ['Hospital_Specialities_Orthopedics', 'Hospital_PhysicianSpeciality_Orthopedics', 'Hospital_PhysicianSpeciality_Orthopedics_Range'], ['Hospital_Specialities_Gynecology', 'Hospital_PhysicianSpeciality_Gynecology', 'Hospital_PhysicianSpeciality_Gynecology_Range'], ['Hospital_Specialities_Obstetrics', 'Hospital_PhysicianSpeciality_Obstetrics', 'Hospital_PhysicianSpeciality_Obstetrics_Range'], ['Hospital_Specialities_Medical_Oncology', 'Hospital_PhysicianSpeciality_Medical_Oncology', 'Hospital_PhysicianSpeciality_Medical_Oncology_Range'], ['Hospital_Specialities_Surgical_Oncology', 'Hospital_PhysicianSpeciality_Surgical_Oncology', 'Hospital_PhysicianSpeciality_Surgical_Oncology_Range'], ['Hospital_Specialities_Radiation_Oncology', 'Hospital_PhysicianSpeciality_Radiation_Oncology', 'Hospital_PhysicianSpeciality_Radiation_Oncology_Range'], ['Hospital_Specialities_Urology', 'Hospital_PhysicianSpeciality_Urology', 'Hospital_PhysicianSpeciality_Urology_Range'], ['Hospital_Specialities_Cardiothoracic_Surgery', 'Hospital_PhysicianSpeciality_Cardiothoracic', 'Hospital_PhysicianSpeciality_Cardiothoracic_Range'], ['Hospital_Specialities_Ent', 'Hospital_PhysicianSpeciality_Ent', 'Hospital_PhysicianSpeciality_Ent_Range'], ['Hospital_Specialities_Endocrinology', 'Hospital_PhysicianSpeciality_Endocrinology', 'Hospital_PhysicianSpeciality_Endocrinology_Range'], ['Hospital_Specialities_Ophthalmology', 'Hospital_PhysicianSpeciality_Ophthalmology', 'Hospital_PhysicianSpeciality_Ophthalmology_Range'], ['Hospital_Specialities_Haematology', 'Hospital_PhysicianSpeciality_Haematology', 'Hospital_PhysicianSpeciality_Haematology_Range'], ['Hospital_Specialities_Rheumatology', 'Hospital_PhysicianSpeciality_Rheumatology', 'Hospital_PhysicianSpeciality_Rheumatology_Range'], ['Hospital_Specialities_Neurology', 'Hospital_PhysicianSpeciality_Neurology', 'Hospital_PhysicianSpeciality_Neurology_Range'], ['Hospital_Specialities_Neuro_Surgery', 'Hospital_PhysicianSpeciality_Neuro_Surgery', 'Hospital_PhysicianSpeciality_Neuro_Surgery_Range'], ['Hospital_Specialities_Plastic_Surgery', 'Hospital_PhysicianSpeciality_Plastic_Surgery', 'Hospital_PhysicianSpeciality_Plastic_Surgery_Range'], ['Hospital_Specialities_Vascular_Surgery', 'Hospital_PhysicianSpeciality_Vascular_Surgery', 'Hospital_PhysicianSpeciality_Vascular_Surgery_Range']]
list_bed_totals = [['Hospital_Beds_General', 'Hospital_Beds_General_Range'], ['Hospital_Beds_Twin_Share', 'Hospital_Beds_Twin_Share_Range'], ['Hospital_Beds_Single', 'Hospital_Beds_Single_Range'], ['Hospital_Beds_ICU_NICU', 'Hospital_Beds_ICU_NICU_Range'], ['Hospital_Beds_Single_AC', 'Hospital_Beds_Single_AC_Range'], ['Hospital_Beds_Twin_Share_AC', 'Hospital_Beds_Twin_Share_AC_Range'], ['Hospital_Beds_Suite', 'Hospital_Beds_Suite_Range'], ['Hospital_Beds_Delux', 'Hospital_Beds_Delux_Range'], ['Hospital_Beds_Others', 'Hospital_Beds_Others_Range'], ['Hospital_Beds_Day_Care', 'Hospital_Beds_Day_Care_Range']]
list_nurse_totals = [['Hospital_NursesGeneral_Ward', 'Hospital_NursesGeneral_Ward_Range'],['Hospital_NursesTwin_Share', 'Hospital_NursesTwin_Share_Range'],['Hospital_NursesSingle', 'Hospital_NursesSingle_Range'],['Hospital_NursesICU', 'Hospital_NursesICU_Range'],['Hospital_NursesSingle_AC', 'Hospital_NursesSingle_AC_Range'],['Hospital_NursesTwin_Share_AC', 'Hospital_NursesTwin_Share_AC_Range'],['Hospital_NursesSuite', 'Hospital_NursesSuite_Range'],['Hospital_Nurses_Delux', 'Hospital_Nurses_Delux_Range'],['Hospital_Nurses_Others', 'Hospital_Nurses_Others_Range'],['Hospital_NursesDay_Care', 'Hospital_NursesDay_Care_Range']]
list_range_totals = ['4-9', '10-20', '21-30', '31-40', '41-50', '51-60', '61-70', '71-80', '81-90', '91-100', '101-200', '201-500', '500<']


root = tk.Tk()
input_filename = askopenfilename(title="Select File to Process...", filetypes=(("Excel Files", "*.xls;*.xlsx"),
                                               ("All files", "*.*") ))
output_filename = asksaveasfilename(title="Export to...", filetypes=(("Excel Files", "*.xls;*.xlsx"),
                                               ("All files", "*.*") ))

if not output_filename.endswith(tuple(['xls', 'xlsx'])):
	output_filename += '.xlsx'

root.withdraw()

if input_filename is not '':
	print('Starting Macro on "' + input_filename + '" (Saving to: ' + output_filename + ')...')
	excel_form = ExcelFormatter(input_filename)
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

	for columns in list_numbers_ranges:
		col_name = " ".join(islice(columns[0].split('_'), 1, None))
		print('\tChecking ' + col_name + ' Ranges ...', end="", flush=True)
		errors_found = excel_form.CheckNumberRangeLines(columns)
		PrintResults(errors_found)


	for columns in list_specialities:
		col_name = " ".join(islice(columns[0].split('_'), 2, None))
		print('\tChecking ' + col_name + ' Specialities ...', end="", flush=True)
		errors_found = excel_form.CheckSpecialitiesLines(columns)
		PrintResults(errors_found)

	print('\tChecking Bed Totals...', end="", flush=True)
	errors_found = excel_form.CheckTotalLines(['Hospital_Beds_Total', 'Hospital_Beds_Total_Range'], list_bed_totals, list_range_totals)
	PrintResults(errors_found)

	print('\tChecking Nurse Totals...', end="", flush=True)
	errors_found = excel_form.CheckTotalLines(['Hospital_NursesTotal', 'Hospital_NursesTotal_Range'], list_nurse_totals, list_range_totals)
	PrintResults(errors_found)

	print('Saving output to ' + output_filename + '...', end="", flush=True)
	excel_form.SaveWorkbook(output_filename)
	print(' Completed!')
	print('Macro Completed!')


