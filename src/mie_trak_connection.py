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

    def get_address_of_party(self, party_fk: int):
        """Gets the address of the party selected using PartyFK"""
        query = " SELECT AddressPK, Name, Address1, Address2, AddressAlt, City, ZipCode FROM Address WHERE PartyFK = ? "
        return self.execute_query(query, party_fk)[0]
    
    def get_state_and_country(self, party_fk):
        """ Fetches the state and country of the selected customer """
        query = " Select StateFK, CountryFK FROM Address where PartyFK = ?"
        info = self.execute_query(query, party_fk)
        state_fk, country_fk = info[0]
        if state_fk:
            state = self.execute_query("Select Description FROM State WHERE StatePK = ?", state_fk)
        else: 
            state = [(None,),] 
        if country_fk:    
            country = self.execute_query("SELECT Description FROM Country WHERE CountryPK = ?", country_fk)
        else:
            country = [(None,),] 
        return state, country

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
                        state,
                        country,
                        customer_rfq_number=None,
                        division_fk=1,
                        received_purchase_order=0,
                        no_bid=0,
                        did_not_get=0,
                        mie_exchange=0,
                        sales_tax_on_freight=0,
                        request_for_quote_status_fk=1, ):
        """ Function for creating RFQ """
        query = ("insert into RequestForQuote (CustomerFK, BillingAddressFK, ShippingAddressFK, DivisionFK, "
                 "ReceivedPurchaseOrder, NoBid, DidNotGet, MIEExchange, SalesTaxOnFreight, RequestForQuoteStatusFK, "
                 "BillingAddressName, BillingAddress1, BillingAddress2, BillingAddressAlt, BillingAddressCity, "
                 "BillingAddressZipCode, ShippingAddressName, ShippingAddress1, ShippingAddress2, ShippingAddressAlt, "
                 "ShippingAddressCity, ShippingAddressZipCode, BillingAddressStateDescription, "
                 "BillingAddressCountryDescription, ShippingAddressStateDescription, "
                 "ShippingAddressCountryDescription, CustomerRequestForQuoteNumber) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,"
                 "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
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
                        state,
                        country,
                        state,
                        country,
                        customer_rfq_number
            ))
            self.cursor.execute("SELECT IDENT_CURRENT('RequestForQuote')")
            rfq_pk = self.cursor.fetchone()[0]
            self.conn.commit()  # Commit the transaction for changes to take effect
        except pyodbc.Error as e:
            print(e)
        return rfq_pk
        
    def upload_documents(self, document_path: str, rfq_fk=None, item_fk=None, document_type_fk=None, secure_document=0, document_group_pk=None, print_with_purchase_order=None):
        """ Function for attaching documents """
        if not rfq_fk and not item_fk:
            raise TypeError("Both values can not be None")
        
        val = rfq_fk if item_fk is None else item_fk
        query = (f"INSERT INTO DOCUMENT (URL, {"RequestForQuoteFK" if item_fk is None else "ItemFk"}, Active, DocumentTypeFK, SecureDocument, DocumentGroupFK, PrintWithPurchaseOrder) VALUES ("
                 f"?,?,?,?,?,?,?)")
        self.cursor.execute(query, (document_path, val, 1, document_type_fk, secure_document, document_group_pk, print_with_purchase_order))
        self.conn.commit()

    def get_or_create_item(self, part_number: str, item_type_fk=1, mps_item=1, purchase=1, forecast_on_mrp=1, mps_on_mrp=1, service_item=1, unit_of_measure_set_fk=1, vendor_unit=1.0, manufactured_item=0, calculation_type_fk=17, inventoriable=1, purchase_order_comment=None,  description=None, comment=None, only_create=None, bulk_ship=1, ship_loose=1):
        """ Fetches ItemPK if the item exists else it creates the item and then returns the ItemPK"""
        if only_create == 1:
            item_inventory_pk = self.create_item_inventory()
            query = ("insert into item (ItemInventoryFk, PartNumber, ItemTypeFK, Description, Comment, MPSItem, "
                     "Purchase, ForecastOnMRP, MPSOnMRP, ServiceItem, PurchaseOrderComment, UnitOfMeasureSetFK, "
                     "VendorUnit, ManufacturedItem, CalculationTypeFK, Inventoriable, BulkShip, ShipLoose) VALUES (?, ?, ?, ?, ?, ?, ?, ?, "
                     "?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
            self.cursor.execute(query, (item_inventory_pk, part_number, item_type_fk, description, comment, mps_item, purchase, forecast_on_mrp, mps_on_mrp, service_item, purchase_order_comment, unit_of_measure_set_fk, vendor_unit, manufactured_item, calculation_type_fk, inventoriable, bulk_ship, ship_loose))
            self.cursor.execute("SELECT IDENT_CURRENT('Item')")
            result = self.cursor.fetchone()[0]
            self.conn.commit()
            return result
        else:
            query = "select ItemPk from item where PartNumber=(?)"
            result = self.cursor.execute(query, part_number).fetchone()
            if result:
                return result[0]
            else:
                item_inventory_pk = self.create_item_inventory()
                query = ("insert into item (ItemInventoryFk, PartNumber, ItemTypeFK, Description, Comment, MPSItem, "
                        "Purchase, ForecastOnMRP, MPSOnMRP, ServiceItem, PurchaseOrderComment, UnitOfMeasureSetFK, "
                        "VendorUnit, ManufacturedItem, CalculationTypeFK, Inventoriable, BulkShip, ShipLoose) VALUES (?, ?, ?, ?, ?, ?, ?, ?, "
                        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
                
                self.cursor.execute(query, (item_inventory_pk, part_number, item_type_fk, description, comment, mps_item, purchase, forecast_on_mrp, mps_on_mrp, service_item, purchase_order_comment, unit_of_measure_set_fk, vendor_unit, manufactured_item, calculation_type_fk, inventoriable, bulk_ship, ship_loose))
                self.cursor.execute("SELECT IDENT_CURRENT('Item')")
                result = self.cursor.fetchone()[0]
                self.conn.commit()
                return result
        
    def create_item_inventory(self):  # HELPER FUNC
        """ Insert data in Item Inventory"""
        self.cursor.execute("insert into ItemInventory (QuantityOnHand) values (0.000)")
        self.cursor.execute("SELECT IDENT_CURRENT('ItemInventory')")
        result = self.cursor.fetchone()[0]
        self.conn.commit()
        return result
    
    def insert_part_details_in_item(self, item_pk, part_number, values, item_type = None):
        """
        Insert values into the Item table where ItemPK = item_pk.
        :param item_pk: The ItemPK to identify the item.
        :param values: A tuple containing the values to be inserted.
        """
        if item_type == 'Material':
            """ """
            po_comment = f" Dimensions (L x W x T): {values[0]} x {values[2]} x {values[1]}"
            try:
                query = "UPDATE Item SET StockLength=?, Thickness=?, StockWidth=?, Weight=?, PartLength=?, PartWidth=?, PurchaseOrderComment=?, ManufacturedItem=?, Purchase=?, ShipLoose=?, BulkShip=? WHERE ItemPK = ?"
                parameters = values[0:4] + (values[0], values[2], po_comment,0,1,0,0, item_pk,)
                self.cursor.execute(query, parameters)
                self.conn.commit()
            except pyodbc.Error as e:
                print(f"Error inserting values into Item table: {e}")
        else:
            try:
                query = "UPDATE Item SET StockLength=?, Thickness=?, StockWidth=?, Weight=?, DrawingNumber=?, DrawingRevision=?, Revision=?, PartLength=?, PartWidth=?, VendorPartNumber=? WHERE ItemPK = ?"
                parameters = values + (values[0], values[2], part_number, item_pk,)
                self.cursor.execute(query, parameters)
                self.conn.commit()
            except pyodbc.Error as e:
                print(f"Error inserting values into Item table: {e}")

    def create_rfq_line_item(self, item_fk: int, request_for_quote_fk: int, line_reference_number: int, quote_fk: int, price_type_fk=3, unit_of_measure_set_fk=1, quantity=None):
        """ Inserts line items in the RFQ """
        insert_query = "INSERT INTO RequestForQuoteLine (ItemFK, RequestForQuoteFK, LineReferenceNumber, QuoteFK, Quantity, PriceTypeFK, UnitOfMeasureSetFK) VALUES (?,?,?,?,?,?,?)"
        self.cursor.execute(insert_query, (item_fk, request_for_quote_fk, line_reference_number, quote_fk, quantity, price_type_fk, unit_of_measure_set_fk))
        self.cursor.execute("SELECT IDENT_CURRENT('RequestForQuoteLine')")
        return_pk = self.cursor.fetchone()[0]
        self.conn.commit()
        return return_pk
       
    def rfq_line_quantity(self, rfq_line_fk, quantity, delivery = 1, price_type_fk = 3):
        """ Inserts the Item quantity in the RFQ Line """
        query = "INSERT INTO RequestForQuoteLineQuantity (RequestForQuoteLineFK, PriceTypeFK, Quantity, Delivery) VALUES (?,?,?,?)"
        self.cursor.execute(query, (rfq_line_fk, price_type_fk, quantity, delivery))
        self.conn.commit()
    
    def create_quote(self, customer_fk, item_fk, quote_type, part_number):
        """ Creates a quote for an item """
        division_fk = 1
        query = "INSERT INTO Quote (CustomerFK, ItemFK, QuoteType, PartNumber, DivisionFK) VALUES (?,?,?,?,?) "
        self.cursor.execute(query, (customer_fk, item_fk, quote_type, part_number, division_fk))
        self.cursor.execute("SELECT IDENT_CURRENT('Quote')")
        return_pk = self.cursor.fetchone()[0]
        self.conn.commit()
        return return_pk
    
    def quote_operations_template(self, quote_fk = None):
        """ Fetches the operation template 494 for the quote """
        self.all_columns = self.get_columns_quote()
        if quote_fk:
            query = f"SELECT {self.all_columns} FROM QuoteAssembly WHERE QuoteFK = {quote_fk}"
            template = self.execute_query(query)
            
        else:
            query = f"SELECT {self.all_columns} FROM QuoteAssembly WHERE QuoteFK = 494 and ItemFk IS NULL "
            template = self.execute_query(query)
        return template

    def get_columns_quote(self):
        """ Fetches all the columns from QuoteAssembly table and converts it into a string """
        query = f'''
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = 'QuoteAssembly' AND COLUMN_NAME NOT IN ('QuoteFK', 'QuoteAssemblyPK', 'LastAccess', 'ParentQuoteAssemblyFK', 'ParentQuoteFK')
        '''
        columns = self.execute_query(query)
        col = [row.COLUMN_NAME for row in columns]
        column_str = ','.join(col)
        return column_str
    
    def quote_operation(self, quote_fk):
        """ Adds the operation template 494 to the quote """
        template = self.quote_operations_template()
        for data in template:
            query = f"INSERT INTO QuoteAssembly (QuoteFK, {self.all_columns}) VALUES ({','.join(['?']*(len(data)+1))})"
            self.cursor.execute(query, (quote_fk,) + tuple(data))
            self.conn.commit()
    
    def get_party_person(self, party_pk):
        """ Gets the customer info to display in the GUI"""
        query = "SELECT ShortName, Email from Party where PartyPK = ?"
        results = self.execute_query(query, party_pk)
        return results[0]
    
    def create_assy_quote(self, quote_to_be_added, quotefk, qty_req = 1):
        query = "INSERT INTO QuoteAssembly (QuoteFK, ItemQuoteFK, SequenceNumber, Pull, Lock, OrderBy, QuantityRequired) VALUES (?,?,?,?,?,?,?)"
        self.cursor.execute(query, (quotefk, quote_to_be_added, 1,0,0,1,qty_req))
        self.cursor.execute("SELECT IDENT_CURRENT('QuoteAssembly')")
        return_pk = self.cursor.fetchone()[0]
        self.conn.commit()
        template = self.quote_operations_template(quote_fk=quote_to_be_added)
        for data in template:
            query = f"INSERT INTO QuoteAssembly (QuoteFK, ParentQuoteAssemblyFK, ParentQuoteFK, {self.all_columns}) VALUES ({','.join(['?']*(len(data)+3))})"
            self.cursor.execute(query, (quotefk, return_pk, quote_to_be_added) + tuple(data))
            self.conn.commit()        
    

    def create_bom_quote(self, quote_fk, item_fk, quote_assembly_seq_number_fk, sequence_number, order_by,
                        party_fk=None, tool=0, stop_sequence=0, unit_of_measure_set_fk=1, setup_time = 0.00, scrape_rebate = 0.000,
                        part_width = 0.00, part_length = 0.000, parts_required = 1.000, quantity_reqd = 1.000, min_piece_price = 0.00,
                        parts_per_blank_scrap_percentage = 0.000, markup_percentage_1 = 9.999999, piece_weight = 0.000, custom_piece_weight = 0.0000, 
                        piece_cost = 0.0000, piece_price = 0.00000, stock_pieces =0, stock_pieces_scrap_perc = 0.000,
                        calculation_type_fk=17, unattended_operation=0, do_not_use_delivery_schedule=0,
                        vendor_unit=1.00000, grain_direction=0, parts_per_blank=1.000,
                        against_grain=0, double_sided=0, cert_reqd=0, non_amortized_item=0,
                        pull=0, not_include_in_piece_price=0, lock=0, nestable=0,
                        bulk_ship=0, ship_loose=0, customer_supplied_material=0):
        """ Adds a BOM to the quote """
        query = "INSERT INTO QuoteAssembly (QuoteFK, ItemFK, PartyFK, UnitofMeasureSetFK, CalculationTypeFK, " \
                "Tool, StopSequence, SequenceNumber, QuoteAssemblySeqNumberFK, UnattendedOperation, " \
                "DoNotUseDeliverySchedule, VendorUnit, GrainDirection, PartsPerBlank, AgainstGrain, " \
                "DoubleSided, CertificationsRequired, NonAmortizedItem, Pull, NotIncludeInPiecePrice, " \
                "Lock, Nestable, BulkShip, ShipLoose, CustomerSuppliedMaterial, OrderBy, SetupTime, ScrapRebate, PartWidth, PartLength, PartsRequired, QuantityRequired, MinimumPiecePrice,  " \
                "PartsPerBlankScrapPercentage, MarkupPercentage1, PieceWeight, CustomPieceWeight, PieceCost, PiecePrice, StockPieces, StockPiecesScrapPercentage) " \
                "VALUES ({})".format(','.join(['?']*41))
        try:
            self.cursor.execute(query, (quote_fk, item_fk, party_fk, unit_of_measure_set_fk,
                                        calculation_type_fk, tool, stop_sequence, sequence_number,
                                        quote_assembly_seq_number_fk, unattended_operation,
                                        do_not_use_delivery_schedule, vendor_unit, grain_direction,
                                        parts_per_blank, against_grain, double_sided, cert_reqd,
                                        non_amortized_item, pull, not_include_in_piece_price, lock,
                                        nestable, bulk_ship, ship_loose, customer_supplied_material, order_by, setup_time, scrape_rebate, part_width, part_length, parts_required, quantity_reqd, min_piece_price,
                                        parts_per_blank_scrap_percentage, markup_percentage_1, piece_weight, custom_piece_weight, piece_cost,
                                        piece_price, stock_pieces, stock_pieces_scrap_perc))
            self.conn.commit()  # Commit the transaction for changes to take effect
        except pyodbc.Error as e:
            print(e)

if __name__ == "__main__":
    m = MieTrak()
