from src.sendmail import send_mail
from src.pull_excel_data import extract_from_excel
import math
from src.mie_trak_connection import MieTrak

conn = MieTrak()
dict1 = {}

def material_for_quote_email(filepath):
    """ returns a dictionary with material as key and its dimensions as values """
    material = extract_from_excel(filepath, "Material")
    length = extract_from_excel(filepath, "Length")
    width = extract_from_excel(filepath, "Width")
    thickness = extract_from_excel(filepath, "Thickness")
    quantity_reqd = extract_from_excel(filepath, "QuantityRequired")
    for a,b,c,d,e in zip(material, length, width, thickness, quantity_reqd):
        a = None if isinstance(a, float) and math.isnan(a) else a
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c
        d = None if isinstance(d, float) and math.isnan(d) else d
        e = None if isinstance(e, float) and math.isnan(e) else e
        if a != None:
            item_inventory_pk = conn.execute_query("Select ItemInventoryFK from Item Where PartNumber = ?", (a,))[0][0]
            quantity_on_hand = conn.execute_query("Select QuantityOnHand from ItemInventory Where ItemInventoryPK = ?", (item_inventory_pk,))[0][0]
            if quantity_on_hand < e:
                dict1[a] = (b,c,d,e)
            else:
                print(f"Material '{a}' has '{quantity_on_hand}' quantity on hand in Item Inventory")
                #TODO: ask user if he still wants to send an email for the quote 
    return dict1

def create_email_body(material_dict):
    " returns an Email body that needs to be send to supplier"
    email_body = "Dear Supplier,\n\n"
    email_body += "We are in need of the following materials and would like to request a quote for each:\n\n"
    
    for material, info in material_dict.items():
        length, width, thickness, quantity = info
        email_body += f"Material: {material}\n"
        email_body += f"Dimensions (Length x Width x Thickness): {length} x {width} x {thickness}\n"
        email_body += f"Quantity Required: {quantity}\n\n"

    email_body += "Please provide us with your best quote at your earliest convenience.\n\n"
    email_body += "Thank you,\nEtezazi Industries"

    return email_body