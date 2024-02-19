# all the mie trak files go here
import pyodbc


class MieTrak:
    def __init__(self):
        # make a connection
        conn_string = "DRIVER={SQL Server};SERVER=ETZ-SQL;DATABASE=SANDBOX;Trusted_Connection=yes"
        self.conn = pyodbc.connect(conn_string)
        self.cursor = self.conn.cursor()

    def execute_query(self, query, parameters=None):
        """Execute a SQL query with optional parameters"""
        try:
            if parameters:
                self.cursor.execute(query, parameters)
            else:
                self.cursor.execute(query)
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            print(f"Error executing query: {query}\nError: {e}")
            return None

    # TODO: also pull billing address, if no billing address then use the same as shipping
    # TODO: make it more robust, raw ideas as of now
    def get_address_of_party(self, party_fk: int):
        """Gets the address of the party selected using PartFK"""
        query = " SELECT AddressPK, Name, Address1, Address2, AddressAlt, City, ZipCode FROM Address WHERE PartyFK = ? "
        return self.execute_query(query, party_fk)[0]

    def insert_into_rfq(self, 
                        customer_fk,
                        billing_address_fk,
                        shipping_address_fk,
                        name, 
                        address1, 
                        address2, 
                        address_alt,
                        city, 
                        zip_code,
                        division_fk=1,
                        received_purchase_order=0,
                        no_bid=0,
                        did_not_get=0,
                        mie_exchange=0,
                        sales_tax_on_freight=0,
                        request_for_quote_status_fk=1,
                        ):
        query = ("insert into RequestForQuote (CustomerFK, BillingAddressFK, ShippingAddressFK, DivisionFK, "
                 "ReceivedPurchaseOrder, NoBid, DidNotGet, MIEExchange, SalesTaxOnFreight, RequestForQuoteStatusFK, "
                 "BillingAddressName, BillingAddress1, BillingAddress2, BillingAddressAlt, BillingAddressCity, "
                 "BillingAddressZipCode, ShippingAddressName, ShippingAddress1, ShippingAddress2, ShippingAddressAlt, "
                 "ShippingAddressCity, ShippingAddressZipCode) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
        try:
            self.cursor.execute(query, (
                        customer_fk,
                        billing_address_fk,
                        shipping_address_fk,
                        division_fk,
                        received_purchase_order,
                        no_bid,
                        did_not_get,
                        mie_exchange,
                        sales_tax_on_freight,
                        request_for_quote_status_fk,
                        name, 
                        address1, 
                        address2, 
                        address_alt,
                        city, 
                        zip_code,
                        name, 
                        address1, 
                        address2, 
                        address_alt,
                        city, 
                        zip_code,
                        
            ))
            self.conn.commit()  # Commit the transaction for changes to take effect
        except pyodbc.Error as e:
            print(e)
        
    def get_rfq_pk(self):
        """Extracting the primary key where the data was inserted in RFQ"""
        query = "SELECT RequestForQuotePK FROM RequestForQuote"
        results = self.execute_query(query)
        return results[-9][0]

    def upload_documents(self, document_path: str, rfq_fk=None, item_fk=None):
        if not rfq_fk and not item_fk:
            raise TypeError("Both values can not be None")
        
        val = rfq_fk if item_fk is None else item_fk
        query = (f"INSERT INTO DOCUMENT (URL, {"RequestForQuoteFK" if item_fk is None else "ItemFk"}, Active) VALUES ("
                 f"?,?,?)")
        self.cursor.execute(query, (document_path, val, 1))
        self.conn.commit()

    def get_or_create_item(self, part_number: str):
        query = "select ItemPk from item where PartNumber=(?)"
        result = self.cursor.execute(query, (part_number)).fetchone()
        if result:
            return result[0]
        else:
            item_inventory_pk = self.create_item_inventory()
            query = "insert into item (ItemInventoryFk, PartNumber, ItemTypeFK) VALUES (?, ?, ?)"
            self.cursor.execute(query, (item_inventory_pk, part_number, 1))
            self.conn.commit()
            result = self.cursor.execute(f"select ItemPK from item where PartNumber = '{part_number}'").fetchall()
            return result[0][0]
        
    def create_item_inventory(self):  # HELPER FUNC
        self.cursor.execute("insert into ItemInventory (QuantityOnHand) values (0.000)")
        self.conn.commit()
        return self.cursor.execute("select ItemInventoryPK from ItemInventory").fetchall()[-1][0]
    
    def create_rfq_line_item(self, item_fk: int, request_for_quote_fk: int, line_reference_number: int, quote_fk: int):
        query = "INSERT INTO RequestForQuoteLine (ItemFK, RequestForQuoteFK, LineReferenceNumber, QuoteFK) VALUES (?,?,?,?)"
        self.cursor.execute(query, (item_fk, request_for_quote_fk, line_reference_number, quote_fk))
        self.conn.commit()
    
    def create_quote(self, customer_fk, item_fk, quote_type, part_number):
        division_fk = 1
        query = "INSERT INTO Quote (CustomerFK, ItemFK, QuoteType, PartNumber, DivisionFK) VALUES (?,?,?,?,?) "
        self.cursor.execute(query, (customer_fk, item_fk, quote_type, part_number, division_fk))
        self.conn.commit()
        return self.cursor.execute("SELECT QuotePK from Quote").fetchall()[-1][0]
    
    def quote_operations_template(self):
        self.all_columns = self.get_columns_quote()
        query = f"SELECT {self.all_columns} FROM QuoteAssembly WHERE QuoteFK = 494 and ItemFk IS NULL "
        template = self.execute_query(query)
        return template
    
    def get_columns_quote(self):
        query = f'''
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = 'QuoteAssembly' AND COLUMN_NAME NOT IN ('QuoteFK', 'QuoteAssemblyPK', 'LastAccess')
        '''
        columns = self.execute_query(query)
        col = [row.COLUMN_NAME for row in columns]
        column_str = ','.join(col)
        return column_str
    
    def quote_operation(self, quote_fk):
        template = self.quote_operations_template()
        for data in template:
            query = f"INSERT INTO QuoteAssembly (QuoteFK, {self.all_columns}) VALUES ({','.join(['?']*(len(data)+1))})"
            self.cursor.execute(query, (quote_fk,) + tuple(data))
            self.conn.commit()
    
    def get_party_person(self, party_pk):
        query = "SELECT ShortName, Email from Party where PartyPK = ?"
        results = self.execute_query(query, party_pk)
        return results[0]
    

    # TODO: Updates for Version2
    # def create_router(self, customer_fk, item_fk):
    #     """ """
    #     division_fk = 1
    #     router_status_fk = 2
    #     router_type = 0
    #     engineering_change = 0
    #     query = "INSERT INTO Router (CustomerFK, ItemFK, DivisionFK, RouterStatusFK, RouterType, EngineeringChange) VALUES (?,?,?,?,?,?)"
    #     self.cursor.execute(query, (customer_fk, item_fk, division_fk, router_status_fk, router_type, engineering_change))
    #     self.conn.commit()
    #     return self.cursor.execute(f"SELECT RouterPK FROM Router WHERE CustomerFK = {customer_fk} AND ItemFK = {item_fk}").fetchall()[-1][0]
    
    # def get_columns_router(self):
    #     query = f'''
    #         SELECT COLUMN_NAME
    #         FROM INFORMATION_SCHEMA.COLUMNS
    #         WHERE TABLE_NAME = 'RouterWorkCenter' AND COLUMN_NAME NOT IN ('RouterFK', 'RouterWorkCenterPK', 'LastAccess')
    #     '''
    #     columns = self.execute_query(query)
    #     col = [row.COLUMN_NAME for row in columns]
    #     column_str = ','.join(col)
    #     return column_str
    
    # def router_work_center_template(self):
    #     self.all_columns_router = self.get_columns_router()
    #     query = f"SELECT {self.all_columns_router} FROM RouterWorkCenter WHERE RouterFK = 4"
    #     template = self.execute_query(query)
    #     return template
    
    # def router_work_center(self, router_fk):
    #     template = self.router_work_center_template()
    #     for data in template:
    #         query = f"INSERT INTO RouterWorkCenter (RouterFK, {self.all_columns_router}) VALUES ({','.join(['?']*(len(data)+1))})"
    #         self.cursor.execute(query, (router_fk,) + tuple(data))
    #         self.conn.commit()


if __name__ == "__main__":
    m = MieTrak()
