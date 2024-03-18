#Functions for getting dictionary. 1st one gives a dictionary with itempk for MAT, FIN, HT
# 2nd one gives dictionary for other information of a part such as Dimensions, Drawing Number
# key for both the dictionary is part number 

from src.pull_excel_data import extract_from_excel
from src.mie_trak_connection import MieTrak
import math

data_base_conn = MieTrak()

def pk_info_dict(filepath):
    part_number = extract_from_excel(filepath, "Part")
    material = extract_from_excel(filepath, "Material")
    finish_code = extract_from_excel(filepath, "FinishCode")
    heat_treat = extract_from_excel(filepath, "HeatTreat")
    length = extract_from_excel(filepath, "Length")
    width = extract_from_excel(filepath, "Width")
    my_dict = {}

    for d, a, b, c, e, f in zip(part_number, material, finish_code, heat_treat, length, width):
        mat_pk = None
        fin_pk = None
        ht_pk = None
        a = None if isinstance(a, float) and math.isnan(a) else a
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c
        e = None if isinstance(e, float) and math.isnan(e) else e
        f = None if isinstance(f, float) and math.isnan(f) else f

        if a:
            result = data_base_conn.execute_query("Select ItemPK from Item where PartNumber = ? and StockLength = ? and StockWidth = ?", (a,e,f))
            if result:
                mat_pk = result[0][0]  # Assuming execute_query returns a list of tuples
            else:
                pk = data_base_conn.get_or_create_item(a, service_item=0, purchase=0, manufactured_item=1, item_type_fk= 2, only_create = 1)
                mat_pk = pk

        if b:
            result = data_base_conn.execute_query("Select ItemPK from Item where PartNumber = ?", (b,))
            if result:
                fin_pk = result[0][0]
            else:
                pk = data_base_conn.get_or_create_item(f"{d} - OP Finish", item_type_fk=5, comment = b, purchase_order_comment=b, inventoriable=0, only_create=1)
                fin_pk = pk

        if c:
            pk = data_base_conn.get_or_create_item(f"{d} - OP HT", item_type_fk=5, description= c, comment= c, purchase_order_comment= c, inventoriable= 0, only_create=1)
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
    pl_revision = extract_from_excel(filepath, "PLRevision")

    info_dict = {}
    for a, b, c, d, e, f, g, h in zip(part_number, length, thickness, width, weight, drawing_number, drawing_revision, pl_revision):
        # Replace NaN values with None
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c
        d = None if isinstance(d, float) and math.isnan(d) else d
        e = None if isinstance(e, float) and math.isnan(e) else e
        f = None if isinstance(f, float) and math.isnan(f) else f
        g = None if isinstance(g, float) and math.isnan(g) else g
        h = None if isinstance(h, float) and math.isnan(h) else h

        info_dict[a] = (b, c, d, e, f, g, h) 
    
    return info_dict
