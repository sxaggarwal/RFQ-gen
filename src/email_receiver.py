from src.pull_excel_data import extract_from_excel
from src.general_class import TableManger
import math
from datetime import datetime

def insert_item_price(filepath, party_fk):
    insert_dict = {
        "PartyFK" : party_fk,
        "ItemFK" : " ",
        "CurrencyCodeFK" : 1,
        "Vendor" : 1,
        "Minimum" : " ", 
        "DivisionFK" : 1,
        "LeadTime" : " ",
        "GoodUntil" : " ",
        "ShippingCharge" : " ",
        "ReceivedDate" : datetime.today().date().strftime('%m/%d/%y')
    }

    item_vendor_table = TableManger("ItemVendor")
    item_table = TableManger("Item")
    item = extract_from_excel(filepath, "Item")
    price = extract_from_excel(filepath, "Price")
    shipping = extract_from_excel(filepath, "ShippingCharge")
    lead_time = extract_from_excel(filepath, "LeadTime")
    good_until = extract_from_excel(filepath, "GoodUntil (MM/DD/YYYY)")

    for a,b,c,d,e in zip(item, price, shipping, lead_time, good_until):
        a = None if isinstance(a, float) and math.isnan(a) else a
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = None if isinstance(c, float) and math.isnan(c) else c
        d = None if isinstance(d, float) and math.isnan(d) else d
        e = None if isinstance(e, float) and math.isnan(e) else e

        if a is not None:
            item_fk = item_table.get("ItemPK", PartNumber=a)[0][0]
            insert_dict["ItemFK"] = item_fk
            insert_dict["Minimum"] = b
            insert_dict["ShippingCharge"] = c
            insert_dict["LeadTime"] = d
            insert_dict["GoodUntil"] = e
            item_vendor_table.insert(insert_dict)   

def get_email_ids(item_type):
    party_buyer_table = TableManger("PartyBuyer")
    party_table = TableManger("Party")
    emails = []
    if item_type == "FIN":
        buyer_fk = party_buyer_table.get("BuyerFK", PartyFK = 3755)
        
    if item_type == "MAT-AL":
        buyer_fk = party_buyer_table.get("BuyerFK", PartyFK = 3744)
    if item_type == "MAT-STEEL":
        buyer_fk = party_buyer_table.get("BuyerFK", PartyFK = 3747)
    
    if item_type == "MAT-EXT":
        buyer_fk = party_buyer_table.get("BuyerFK", PartyFK = 3751)

    if item_type == "HT-AL":
        buyer_fk = party_buyer_table.get("BuyerFk", PartyFK = 3764, Description = "ALUMINUM HT")
    if item_type == "HT-STEEL":
        buyer_fk = party_buyer_table.get("BuyerFk", PartyFK = 3764, Description = "STEEL HT")
    if item_type == "HT-STEEL" or "HT-AL":
        buyer_fk_1 = party_buyer_table.get("BuyerFk", PartyFK = 3764, Description = "STEEL/ALUMINUM HT")
        if buyer_fk_1:
            for fk in buyer_fk_1:
                email = party_table.get("Email", PartyPK = fk[0])[0][0]
                emails.append(email)
    
    for fk in buyer_fk:
        email = party_table.get("Email", PartyPK = fk[0])[0][0]
        emails.append(email)

    return emails

def get_item_type(quote_fk = None, itemfk = None):
    quote_assembly_table = TableManger("QuoteAssembly")
    item_table = TableManger("Item")
    mat_email = []
    ht_email = []
    fin_email = []
    if quote_fk:
        itemfk = quote_assembly_table.get("ItemFK", QuoteFK=quote_fk)
        for fk in itemfk:
            if fk[0] is not None:
                description = item_table.get("Description", ItemPK=fk[0])[0][0]
                if description:
                    description = description.lower()
                    if 'al' in description or 'aluminum' in description:
                        ht_email.extend(get_email_ids("HT-AL"))
                        mat_email.extend(get_email_ids("MAT-AL"))
                        fin_email.extend(get_email_ids("FIN"))
                        break
                    elif 'steel' in description or 'st' in description:
                        ht_email.extend(get_email_ids("HT-STEEL"))
                        mat_email.extend(get_email_ids("MAT-STEEL"))
                        fin_email.extend(get_email_ids("FIN"))
                        break
    else:
        description = item_table.get("Description", ItemPK=itemfk)[0][0]
        if description:
            description = description.lower()
            if 'al' in description or 'aluminum' in description:
                ht_email.extend(get_email_ids("HT-AL"))
                mat_email.extend(get_email_ids("MAT-AL"))
                fin_email.extend(get_email_ids("FIN"))

            elif 'steel' in description or 'st' in description:
                ht_email.extend(get_email_ids("HT-STEEL"))
                mat_email.extend(get_email_ids("MAT-STEEL"))
                fin_email.extend(get_email_ids("FIN"))
                


    return mat_email, ht_email, fin_email


if __name__ == "__main__":
    filepath = r"c:\Users\saggarwal\Documents\quote_price.xlsx"
    # insert_item_price(filepath, 2167)
    mat_email, ht_email, fin_email = get_item_type(quote_fk=3060)
    print(f"HT_Email: {ht_email}")
    print(f"FIN_Email: {fin_email}")
    print(f"Mat_Email: {mat_email}")



