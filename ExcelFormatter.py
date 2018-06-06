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
						cell.fill = PatternFill(start_color=self.color_red, end_color=self.color_red, fill_type='solid')
						errors += 1
					elif not cell.value.strip().endswith(','):
						cell.fill = PatternFill(start_color=self.color_yellow, end_color=self.color_yellow, fill_type='solid')
						cell.value = cell.value.strip() + ','
						errors += 1
					else:
						cell.value = cell.value.strip()
		return errors

	def CheckEmailAddressLines(self, column_names):
		errors = 0
		for idx in self.GetColumnIndicesByName(column_names):
			for index, selected_column in enumerate(islice(self.ws.iter_cols(), idx, idx+1)):
				for cell in islice(selected_column, 2, None):
					if cell.value is None:
						cell.fill = PatternFill(start_color=self.color_red, end_color=self.color_red, fill_type='solid')
						errors += 1
					elif cell.value.strip() == "NI" or cell.value.strip() == "NA" or cell.value.strip() == "N0"  or cell.value.strip() == "NP":
						cell.value = "NO"
					elif not re.match(r"[^@]+@[^@]+\.[^@]+", cell.value.strip()) and cell.value.strip() != "NO":
						cell.fill = PatternFill(start_color=self.color_yellow, end_color=self.color_yellow, fill_type='solid')
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
						cell.fill = PatternFill(start_color=self.color_red, end_color=self.color_red, fill_type='solid')
						errors += 1
					elif cell.value.strip() == "NI" or cell.value.strip() == "NA" or cell.value.strip() == "N0"  or cell.value.strip() == "NP":
						cell.value = "NO"
					elif not re.match(r"^(\+\d{1,3}\ \d{10})$", cell.value.strip()) and cell.value.strip() != "NO":
						cell.fill = PatternFill(start_color=self.color_yellow, end_color=self.color_yellow, fill_type='solid')
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
						cell.fill = PatternFill(start_color=self.color_red, end_color=self.color_red, fill_type='solid')
						errors += 1
					elif not isinstance(cell.value, float) and (cell.value.strip() == "NO" or cell.value.strip() == "NA" or cell.value.strip() == "N0"  or cell.value.strip() == "NP"):
						cell.value = "NI"
					elif not re.match(r"^(\d{1,2}.\d{4,10} \d{1,3}.\d{4,10})$", str(cell.value)) and str(cell.value) != "NI":
						cell.fill = PatternFill(start_color=self.color_yellow, end_color=self.color_yellow, fill_type='solid')
						errors += 1
		return errors