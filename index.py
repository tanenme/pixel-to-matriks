import numpy as np
import cv2
from docx import Document

# Membaca gambar menggunakan OpenCV
# Menggantikan "gambar.jpg" dengan nama file gambar Anda
image = cv2.imread(
    r"./input.jpg", cv2.IMREAD_COLOR)

# Mengubah gambar menjadi matriks NumPy
matrix = np.array(image)

# Menyusun matriks menjadi ukuran yang diinginkan
desired_rows = 175
desired_cols = 200
matrix = cv2.resize(matrix, (desired_cols, desired_rows))

output_file = "output.docx"  # Nama file output yang diinginkan
doc = Document()
for row in matrix:
    doc.add_paragraph(" ".join(str(cell) for cell in row))
doc.save(output_file)