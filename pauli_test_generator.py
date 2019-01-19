from docx import Document
from docx.shared import Inches 
import random

jumlah_hari = 30

for i in range(jumlah_hari):
	print("######")
	print("membuat document ke-" + str(i + 1) + " ...")

	document = Document()

	# set ukuran kertas menjadi A4
	section = document.sections[0]
	section.page_height = Inches(11.69)
	section.page_width = Inches(8.27)

	# set banyak halaman
	row_per_page = 30
	page_per_hari = 20

	total_row = page_per_hari * row_per_page
	total_column = 14

	table = document.add_table(rows=total_row, cols=total_column)

	page_counter = 0
	for row in table.rows:
		if page_counter % 30 == 0:
			print("halaman ke-", (page_counter // 30))
		page_counter += 1
		for cell in row.cells:
			cell.text = str(random.randint(0,9))

	# saving the document
	document.save("Pauli Test_by uul_" + str(i + 1) +".docx")
	print("Saving Pauli Test_by uul_" + str(i + 1) +".docx .....\n\n") 