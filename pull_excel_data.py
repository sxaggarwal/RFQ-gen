import pandas as pd


def extract_from_excel(filepath, column_name):
    """ Extracting a column from an excel file"""
    df = pd.read_excel(filepath)
    part_numbers = df[column_name].tolist()
    return part_numbers