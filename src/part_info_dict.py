from src.pull_excel_data import extract_from_excel
from src.Mie_trak_connection import MieTrak

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
