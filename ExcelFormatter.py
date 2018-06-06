import openpyxl as oxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from itertools import islice
import re

class ExcelFormatter(object):
	def __init__(self, filename):
		self.filename = filename
		self.wb = oxl.load_workbook(filename)
		self.ws = self.wb.active
		self.ProcessColumnIndices()

		self.color_red = 'FFFF0000'
		self.color_yellow = 'FFFFFF00'

	def SaveWorkbook(self, new_filename):
		self.wb.save(new_filename)

	def ProcessColumnIndices(self):
		self.indices = {}
		n = 0
		for ix, first_row in enumerate(islice(self.ws.iter_rows(), 1)):
			for cell in first_row:
				self.indices[cell.value] = n
				n += 1

	def GetColumnIndicesByName(self, column_names):
		indices = []
		for c_name in column_names:
			indices.append(self.indices.get(c_name))
		return indices

	def CheckAddressLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif not cell.value.strip().endswith(','):
						cell.value = cell.value.strip() + ','
					else:
						cell.value = cell.value.strip()
		return errors

	def CheckEmailAddressLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif self.isNotAvailable(cell.value.strip()):
						cell.value = "NO"
					elif not re.match(r"[^@]+@[^@]+\.[^@]+", cell.value.strip()) and cell.value.strip() != "NO":
						self.fillCellColour(cell, self.color_yellow)
						errors += 1
					else:
						cell.value = cell.value.strip()
		return errors

	def CheckPhoneNumbersLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif self.isNotAvailable(cell.value.strip()):
						cell.value = "NO"
					elif not re.match(r"^(\+\d{1,3}\ \d{10})$", cell.value.strip()) and cell.value.strip() != "NO":
						# Replacing invisible non-really-a-space
						new_str = str(cell.value).replace(' ', '-')
						p = re.compile(r"^(\+\d{1,3}\-\d{10})$")
						found_val = p.search(new_str)
						if found_val is not None:
							cell.value = found_val.group(0).replace('-', ' ')
						else:
							self.fillCellColour(cell, self.color_yellow)
							errors += 1
					else:
						cell.value = cell.value.strip()
		return errors

	def CheckTitleLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif not (cell.value.strip() in ["Mr.", "Dr.", "Mrs.", "Ms.", "Adv."]):
						if "mrs" in cell.value.strip().lower():
							cell.value = "Mrs."
						elif "mr" in cell.value.strip().lower():
							cell.value = "Mr."
						elif "ms" in cell.value.strip().lower():
							cell.value = "Ms."
						elif "dr" in cell.value.strip().lower():
							cell.value = "Dr."
						elif "adv" in cell.value.strip().lower():
							cell.value = "Adv."
						else:
							self.fillCellColour(cell, self.color_yellow)
							errors += 1
					else:
						cell.value = cell.value.strip()
		return errors

	def CheckLatLonLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif not isinstance(cell.value, float) and self.isNotAvailable(cell.value.strip()):
						cell.value = "NI"
					elif not re.match(r"^(\d{1,2}.\d{4,10} \d{1,3}.\d{4,10})$", str(cell.value)) and str(cell.value) != "NI":
						self.fillCellColour(cell, self.color_yellow)
						errors += 1
		return errors

	def CheckHoursLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif not re.match(r"^([01]?[0-9]|2[0-3]):[0-5][0-9]$", str(cell.value)):
						p = re.compile(r"([01]?[0-9]|2[0-3]):[0-5][0-9]")
						found_val = p.search(str(cell.value))
						if found_val is not None:
							cell.value = found_val.group(0)
						else:
							self.fillCellColour(cell, self.color_yellow)
							errors += 1
		return errors

	# def CheckCountOrRanger(self, column_names):
	# 	errors = 0
	# 	for idx in self.GetColumnIndicesByName(column_names):
	# 		for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
	# 			for cell in islice(selected_column, 2, None):
	# 				if cell.value is None:
	# 					# check adjacent row

	# 					# if adjacent row is empty and not "NA"
	# 						 # fill both cells
	# 						self.fillCellColour(cell, self.color_red)
	# 						errors += 1
	# 				elif self.isNotAvailable(cell.value.strip()):
	# 					cell.value = "NA"
	# 				elif not re.match(r"^(\d{1,5})$", str(cell.value)) and str(cell.value) != "NA":
	# 					self.fillCellColour(cell, self.color_yellow)
	# 					errors += 1
	# 	return errors

	def CheckYYNOAPPLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif cell.value.strip().lower() == "yy" or cell.value.strip().lower() == "app":
						cell.value = "YY"
					elif cell.value.strip().lower() == "no":
						cell.value = "NO"
					else:
						self.fillCellColour(cell, self.color_yellow)
						errors += 1
		return errors

	def CheckAllowedName(self, column_names, state_names):
		errors = 0
		cities = []
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						self.fillCellColour(cell, self.color_red)
						errors += 1
					elif cell.value.strip() in state_names:
						cell.value = cell.value.strip()
					else:
						self.fillCellColour(cell, self.color_yellow)
						errors += 1
		return errors

	def fillCellColour(self, cell, colour):
		cell.fill = PatternFill(start_color=colour, end_color=colour, fill_type='solid')

	def isNotAvailable(self, value):
		return value == "NO" or value == "NI" or value == "NA" or value == "N0"  or value == "NP"