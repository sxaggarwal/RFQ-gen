#Functions for getting dictionary. 1st one gives a dictionary with itempk for MAT, FIN, HT
# 2nd one gives dictionary for other information of a part such as Dimensions, Drawing Number
# key for both the dictionary is part number 

from src.pull_excel_data import extract_from_excel
from src.Mie_trak_connection import MieTrak
import math

data_base_conn = MieTrak()

def pk_info_dict(filepath):
    part_number = extract_from_excel(filepath, "Part")
    material = extract_from_excel(filepath, "Material")
    finish_code = extract_from_excel(filepath, "FinishCode")
    heat_treat = extract_from_excel(filepath, "HeatTreat (Y/N)")
    my_dict = {}

    for d, a, b, c in zip(part_number, material, finish_code, heat_treat):
        mat_pk = None
        fin_pk = None
        ht_pk = None
        a = None if isinstance(a, float) and math.isnan(a) else a
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c

        if a:
            result = data_base_conn.execute_query("Select ItemPK from Item where PartNumber = ?", (a,))
            if result:
                mat_pk = result[0][0]  # Assuming execute_query returns a list of tuples
            else:
                pk = data_base_conn.get_or_create_item(a)
                mat_pk = pk

        if b:
            result = data_base_conn.execute_query("Select ItemPK from Item where PartNumber = ?", (b,))
            if result:
                fin_pk = result[0][0]
            else:
                pk = data_base_conn.get_or_create_item(f"OP Finish - {d}")
                fin_pk = pk

        if c and (c == 'Y' or c == 'y'):
            pk = data_base_conn.get_or_create_item(f"OP HT - {d}")
            ht_pk = pk

        my_dict[d] = (mat_pk, ht_pk, fin_pk)

    return my_dict

def part_info(filepath):
    part_number = extract_from_excel(filepath, "Part")
    length = extract_from_excel(filepath, "Length")
    thickness = extract_from_excel(filepath, "Thickness")
    width = extract_from_excel(filepath, "Width")
    weight = extract_from_excel(filepath, "Weight")
    drawing_number = extract_from_excel(filepath, "DrawingNumber")
    drawing_revision = extract_from_excel(filepath, "DrawingRevision")

    info_dict = {}
    for a, b, c, d, e, f, g in zip(part_number, length, thickness, width, weight, drawing_number, drawing_revision):
        # Replace NaN values with None
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c
        d = None if isinstance(d, float) and math.isnan(d) else d
        e = None if isinstance(e, float) and math.isnan(e) else e
        f = None if isinstance(f, float) and math.isnan(f) else f
        g = None if isinstance(g, float) and math.isnan(g) else g
        
        info_dict[a] = (b, c, d, e, f, g) 
    
    return info_dict
