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

	def CheckNumberRangeLines(self, column_names):
		errors = 0
		column_number_idx = self.GetColumnIndicesByName(column_names)[0]
		column_range_idx = self.GetColumnIndicesByName(column_names)[1]
		cell_index = 3
		for index, selected_column in enumerate(islice(self.ws.iter_cols(), column_number_idx, column_number_idx+1)):
			for cell in islice(selected_column, 2, None):
				range_cell = self.ws.cell(row=cell_index, column=column_range_idx + 1)
				errors += self.CheckCountOrRange(cell, range_cell)
				cell_index += 1
		return errors

	def CheckCountOrRange(self, cell_number, cell_range):
		if cell_number.value is None:
			if cell_range.value is None:
				self.fillCellColour(cell_number, self.color_red)
				self.fillCellColour(cell_range, self.color_red)
				return 1
			else:
				if not re.match(r"^[']?\d{1,10}-\d{1,10}$", str(cell_range.value).strip()) and str(cell_range.value).strip() != "NA":
					self.fillCellColour(cell_range, self.color_yellow)
					return 1
		elif not re.match(r"^\d{1,10}$", str(cell_number.value).strip()):
			self.fillCellColour(cell_number, self.color_yellow)
			return 1
		return 0

	def CheckSpecialitiesLines(self, column_names):
		errors = 0
		column_spec_idx = self.GetColumnIndicesByName(column_names)[0]
		column_number_idx = self.GetColumnIndicesByName(column_names)[1]
		column_range_idx = self.GetColumnIndicesByName(column_names)[2]
		cell_index = 3
		for index, selected_column in enumerate(islice(self.ws.iter_cols(), column_spec_idx, column_spec_idx+1)):
			for cell in islice(selected_column, 2, None):
				number_cell = self.ws.cell(row=cell_index, column=column_number_idx + 1)
				range_cell = self.ws.cell(row=cell_index, column=column_range_idx + 1)
				if str(cell.value).strip().upper() == "YY":
					cell.value = "YY"
					errors += self.CheckCountOrRange(number_cell, range_cell)
				elif str(cell.value).strip().upper() == "NA" or str(cell.value).strip().upper() == "NO":
					cell.value = "NA"
					if number_cell.value is not None:
						self.fillCellColour(number_cell, self.color_red)
						errors += 1
					if str(range_cell.value).strip() != "NA":
						self.fillCellColour(range_cell, self.color_red)
						errors += 1
				else:
					self.fillCellColour(cell, self.color_red)
				cell_index += 1
		return errors

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