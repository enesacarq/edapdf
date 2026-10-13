
from pathlib import Path
import fitz

Path("results").mkdir(exist_ok=True)

document = fitz.open()
page = document.new_page()

page.insert_text(
    (50, 50),
    "Student List",
    fontsize=16
)

# Titujt e tabeles
headers = ["Name", "Age", "Grade"]
rows = [
    ["Ali", "22", "85"],
    ["Ayse", "24", "90"],
    ["Mehmet", "21", "78"]
]

x_positions = [50, 200, 350]
y_start = 100
row_height = 35
column_width = 150

# Vizatojme tabelen
for row_index in range(5):
    y = y_start + row_index * row_height
    page.draw_line(
        (50, y),
        (500, y)
    )

for column_index in range(4):
    x = 50 + column_index * column_width
    page.draw_line(
        (x, y_start),
        (x, y_start + 4 * row_height)
    )

# Shkruajme titujt
for index, header in enumerate(headers):
    page.insert_text(
        (x_positions[index] + 10, y_start + 23),
        header,
        fontsize=12
    )

# Shkruajme te dhenat
for row_index, row in enumerate(rows):
    for column_index, value in enumerate(row):
        page.insert_text(
            (
                x_positions[column_index] + 10,
                y_start + (row_index + 1) * row_height + 23
            ),
            value,
            fontsize=12
        )

document.save("results/test_table.pdf")
document.close()

print("PDF me tabele u krijua me sukses!")
