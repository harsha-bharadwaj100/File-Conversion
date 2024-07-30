import os
import sys
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
from openpyxl import load_workbook
from pptx import Presentation
import pandas as pd


def pdf_to_docx(input_file, output_file):
    cv = Converter(input_file)
    cv.convert(output_file)
    cv.close()


def docx_to_pdf(input_file, output_file):
    convert(input_file, output_file)


def image_converter(input_file, output_file):
    img = Image.open(input_file)
    img.save(output_file)


def excel_to_csv(input_file, output_file):
    df = pd.read_excel(input_file)
    df.to_csv(output_file, index=False)


def csv_to_excel(input_file, output_file):
    df = pd.read_csv(input_file)
    df.to_excel(output_file, index=False)


def ppt_to_pdf(input_file, output_file):
    prs = Presentation(input_file)
    prs.save(output_file)


def main():
    if len(sys.argv) != 4:
        print(
            "Usage: python file_converter.py <input_file> <output_file> <conversion_type>"
        )
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    conversion_type = sys.argv[3].lower()

    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' does not exist.")
        sys.exit(1)

    conversion_functions = {
        "pdf2docx": pdf_to_docx,
        "docx2pdf": docx_to_pdf,
        "jpg2png": image_converter,
        "png2jpg": image_converter,
        "excel2csv": excel_to_csv,
        "csv2excel": csv_to_excel,
        "ppt2pdf": ppt_to_pdf,
        "img2pdf": image_converter,
    }

    if conversion_type not in conversion_functions:
        print(f"Error: Unsupported conversion type '{conversion_type}'.")
        print("Supported conversions:", ", ".join(conversion_functions.keys()))
        sys.exit(1)

    try:
        conversion_functions[conversion_type](input_file, output_file)
        print(f"Conversion completed: {input_file} -> {output_file}")
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
