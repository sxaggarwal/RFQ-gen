import fitz, pathlib, logging

FILEPATH = r"Y:\Estimating\Non-restricted\SPIRIT\RFE 141776P\ENGINEERING\313W3153-7.pdf"


def boeing_pdf_converter(input_filepath, output_filename):  #TODO: args taken from the terminal
    """PDF -> txt file: maintains the correct format of the pdf"""
    with fitz.open(input_filepath) as doc:
        print(doc[1].get_text())
        text = chr(12).join([page.get_text() for page in doc])

    pathlib.Path(output_filename).write_bytes(text.encode())
    return output_filename + ".txt"


def get_headers(parsed_text_file):
    """Does not return anything, just added headers on to the final dict. Uses string methods"""
    with open(parsed_text_file, "r") as doc:
        data_list = [line for line in doc.readlines()]
    for line in data_list:
        if line.strip().startswith("PARTS LIST"):  # this can be done only once using regex matching
            header_list = line.split()
            headers = {"Date": header_list[3],
                        "PL-Number": header_list[5],
                        "Rev": header_list[6]}
            return headers

if __name__ == "__main__":
    # boeing_pdf_converter(r"M:\docparser\testing_pdfs\PSE_112T1750_PL__AC_COM___PROPER_1_TO_135.pdf", "test_raw_file1")
    boeing_pdf_converter(r"Y:/PDM/Non-restricted/BOEING/B02-20289-1/B02-20289_PL_Sht+1.pdf", "test_raw_file2")
    