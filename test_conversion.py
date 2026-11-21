
from pdf_processing import pdf_to_images

image_paths = pdf_to_images(
    input_path="results/test_table.pdf",
    output_dir="output_images",
    image_format="png"
)

print("Konvertimi përfundoi me sukses!")
print("Imazhet e krijuara:", image_paths)


from pdf_processing import images_to_pdf

pdf_path = images_to_pdf(
    input_paths=["output_images/page_1.png"],
    output_path="results/converted.pdf"
)

print("PDF u krijua me sukses!")
print("Vendndodhja:", pdf_path)



from pdf_processing import merge_pdfs

merged_path = merge_pdfs(
    input_paths=[
        "results/test_table.pdf",
        "results/converted.pdf"
    ],
    output_path="results/merged.pdf"
)

print("PDF-të u bashkuan me sukses!")
print("Dokumenti i krijuar:", merged_path)



from pdf_processing import split_pdf

split_files = split_pdf(
    input_path="results/merged.pdf",
    page_ranges=[(1, 1), (2, 2)],
    output_dir="results/split"
)

print("PDF u nda me sukses!")
print("Dokumentet e krijuara:", split_files)



from pdf_processing import delete_pdf_pages

deleted_path = delete_pdf_pages(
    input_path="results/merged.pdf",
    pages=[2],
    output_path="results/deleted_pages.pdf"
)

print("Faqja u fshi me sukses!")
print("Dokumenti i krijuar:", deleted_path)



from pdf_processing import add_watermark

watermarked_path = add_watermark(
    input_path="results/merged.pdf",
    output_path="results/watermarked.pdf",
    text="DRAFT"
)

print("Watermark u shtua me sukses!")
print("Dokumenti i krijuar:", watermarked_path)



from pdf_processing import pdf_to_word

word_path = pdf_to_word(
    input_path="results/test_table.pdf",
    output_path="results/converted.docx"
)

print("PDF u konvertua ne Word me sukses!")
print("Dokumenti i krijuar:", word_path)



from pdf_processing import pdf_to_powerpoint

powerpoint_path = pdf_to_powerpoint(
    input_path="results/merged.pdf",
    output_path="results/converted.pptx"
)

print("PDF u konvertua ne PowerPoint me sukses!")
print("Dokumenti i krijuar:", powerpoint_path)



from pdf_processing import pdf_to_excel

excel_path = pdf_to_excel(
    input_path="results/test_table.pdf",
    output_path="results/converted.xlsx"
)

print("PDF u konvertua ne Excel me sukses!")
print("Dokumenti i krijuar:", excel_path)



from pdf_processing import office_to_pdf

# Test 1: Word to PDF
word_pdf = office_to_pdf(
    input_path="results/converted.docx",
    output_path="results/word_to_pdf.pdf"
)

print("Word u konvertua ne PDF:", word_pdf)


# Test 2: Excel to PDF
excel_pdf = office_to_pdf(
    input_path="results/converted.xlsx",
    output_path="results/excel_to_pdf.pdf"
)

print("Excel u konvertua ne PDF:", excel_pdf)


# Test 3: PowerPoint to PDF
powerpoint_pdf = office_to_pdf(
    input_path="results/converted.pptx",
    output_path="results/powerpoint_to_pdf.pdf"
)

print("PowerPoint u konvertua ne PDF:", powerpoint_pdf)
