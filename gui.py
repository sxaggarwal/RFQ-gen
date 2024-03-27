# all tkinter modules that go on the screen
from src.mie_trak_connection import MieTrak
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from src.pull_excel_data import extract_from_excel
import os
import shutil
import math
from src.part_info_dict import pk_info_dict, part_info
from src.sendmail import send_mail
from src.emailer import material_for_quote_email, create_email_body
from src.email_receiver import get_email_ids


class RfqGen(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RFQGen")
        self.geometry("305x400")

        self.data_base_conn = MieTrak() 
        self.customer_names = [customer[0] for customer in self.data_base_conn.execute_query("select name from party")]
        query = "select partypk, name from party"
        self.data = self.data_base_conn.execute_query(query)
        self.customer_names = [d[1] for d in self.data]
        self.customer_number_to_partypk = {i: d[0] for i, d in enumerate(self.data, start=1)}

        self.make_combobox()

    def make_combobox(self):

        # Customer select combobox
        tk.Label(self, text="Select Customer: ").grid(row=0, column=0)
        self.customer_select_box = ttk.Combobox(self, values=self.customer_names, state="readonly")
        self.customer_select_box.grid(row=1, column=0)
        
        tk.Label(self, text="Selected Customer Info: ").grid(row=2, column=0)
        self.customer_info_text = tk.Text(self, height=4, width=30)
        self.customer_info_text.grid(row=3, column=0)

        tk.Label(self, text="Enter RFQ Number: ").grid(row=4, column=0)
        self.rfq_number_text = tk.Entry(self, width=20)
        self.rfq_number_text.grid(row=5, column=0)

        # Bind the combobox selection event to update customer information
        self.customer_select_box.bind("<<ComboboxSelected>>", self.update_customer_info)

        # Entrybox for the requested parts in an Excel file. (upload for Excel file)
        tk.Label(self, text="Parts Requested File:").grid(row=6, column=0)
        self.file_path_PR_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PR_entry.grid(row=7, column=0)

        # this should be at the bottom
        browse_button_1 = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("Excel files", self.file_path_PR_entry))
        browse_button_1.grid(row=8, column=0)

        # Selection/ Upload for PartList
        tk.Label(self, text="Part Lists File (PL):").grid(row=9, column=0)
        self.file_path_PL_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PL_entry.grid(row=10, column=0)

        browse_button_part_list = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("All files", self.file_path_PL_entry))
        browse_button_part_list.grid(row=11, column=0)

        # Checkbox for ITAR RESTRICTED
        self.itar_restricted_var = tk.BooleanVar()
        self.itar_restricted_checkbox = tk.Checkbutton(self, text="ITAR RESTRICTED", variable=self.itar_restricted_var)
        self.itar_restricted_checkbox.grid(row=12, column=0)

        # main button
        generate_button = tk.Button(self, text="Generate RFQ", command=self.generate_rfq)
        generate_button.grid(row=13, column=0)

        sending_email_button = tk.Button(self, text="Send Email", command=self.sending_email)
        sending_email_button.grid(row=14, column=0)
    
    def update_customer_info(self, event=None):
        """Update customer information label when a customer is selected."""
        info = self.get_party_person_info()
        self.customer_info_text.delete(1.0, tk.END)
        if info:
            short_name, email = info[0]
            self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
        else:
            self.customer_info_text.insert(tk.END, "No information available")

    def get_party_pk(self):
        """ Get the selected customer's partypk """
        selected_customer_index = self.customer_select_box.current()
        return self.customer_number_to_partypk[selected_customer_index+1]
    
    def get_party_person_info(self):
        """ Retrieve short name and email for the selected customer """
        self.selected_customer_partypk = self.get_party_pk()
        query = "SELECT ShortName, Email FROM Party WHERE PartyPK = ?"
        info = self.data_base_conn.execute_query(query, self.selected_customer_partypk)
        return info
        
    def browse_files_parts_requested(self, filetype: str, list_box):
        """ Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [filepath for filepath in filedialog.askopenfilenames(title="Select Files", filetypes=(param,))]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror("File Browse Error", "An error occurred during file selection. Please try again.")

    @staticmethod
    def transfer_file_to_folder(folder_path: str, file_path: str) -> str:
        """ Copies file from one path to another path"""
        os.makedirs(folder_path, exist_ok=True)

        filename = os.path.basename(file_path)  # source file path
        destination_path = os.path.join(folder_path, filename)
        shutil.copyfile(file_path, destination_path)

        return destination_path
    
    def show_selection(self):
        """ This will display a selection box for which we need to send email """
        items = ["MAT-AL", "MAT-STEEL", "MAT-EXT", "FIN", "HT-AL", "HT-STEEL" ]
        popup = tk.Toplevel(self)
        popup.title("Select Item")
        popup.geometry("300x150")
        popup.grab_set()
        tk.Label(popup, text="Select Item: ").pack()
        item_select_box = ttk.Combobox(popup, values= items, state="readonly")
        item_select_box.pack()

        def send_email():
            selected_item = item_select_box.get()
            if not selected_item:
                messagebox.showerror("Error", "Please select an item.")
                return
            
            email_body = create_email_body(material_for_quote_email(self.file_path_PR_entry.get(0, tk.END)[0]), selected_item)
            self.show_email_body_popup(email_body, selected_item)
            popup.destroy()
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)

        send_button = tk.Button(popup, text="Send", command=send_email)
        send_button.pack()


    def show_email_body_popup(self, email_body, item_type):
        """Popup to display the body of the email that will be sent with Edit and Confirm options"""
        popup = tk.Toplevel(self)
        popup.title("Email Body")
        popup.grab_set()

        # Create a Text widget for displaying the email body
        body_text = tk.Text(popup, wrap=tk.WORD, height=20, width=60)
        body_text.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Insert the email body into the Text widget
        body_text.insert(tk.END, email_body)
        body_text.config(state="disabled")

        # Create a scrollbar and associate it with the Text widget
        scrollbar = tk.Scrollbar(popup, command=body_text.yview)
        scrollbar.grid(row=0, column=2, sticky="ns")
        body_text.config(yscrollcommand=scrollbar.set)

        # Create Edit and Confirm buttons
        edit_button = tk.Button(popup, text="Edit", command=lambda: self.edit_email_body(email_body, popup, item_type))
        edit_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        confirm_button = tk.Button(popup, text="Confirm", command=lambda: self.confirm_send_email(email_body, popup, item_type))
        confirm_button.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # Make the Text widget and scrollbar expandable
        popup.rowconfigure(0, weight=1)
        popup.columnconfigure(0, weight=1)

    def edit_email_body(self, email_body, popup, item_type):
        """ If user wants to edit the email body then this will show up """
        edit_window = tk.Toplevel(popup)
        edit_window.title("Edit Email Body")
        edit_window.grab_set()
        
        edit_text = tk.Text(edit_window, width=100, height=50)
        edit_text.insert(tk.END, email_body)
        edit_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = tk.Scrollbar(edit_window, command=edit_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        edit_text.config(yscrollcommand=scrollbar.set)

        save_button = tk.Button(edit_window, text="Save", command=lambda: self.save_email_body(edit_text, popup, item_type))
        save_button.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky="nsew")

    def save_email_body(self, edit_text, popup, item_type):
        """ Saves the new body if user edits the email"""
        new_body = edit_text.get("1.0", tk.END)
        self.show_email_body_popup(new_body, item_type)
        popup.destroy()

    def confirm_send_email(self, email_body, popup, item_type):
        """ This will popup a confirmation box if the user clicks on confirm"""
        popup.destroy()
        
        email_ids = []
        ids = get_email_ids(item_type)
        
        for email in ids:
            # Check if email is a float and NaN
            if isinstance(email, float) and math.isnan(email):
                email = None
            email_ids.append(email)
            
        # Remove None values
        email_ids = [email for email in email_ids if email is not None]
        
        confirmation = messagebox.askyesno("Confirm", "Are you sure you want to send this email?")
        
        if confirmation:
            self.edit_email_ids(email_body, email_ids)
        else:
            self.show_email_body_popup(email_body, item_type)
    
    def edit_email_ids(self, email_body, email_ids):
        """Popup to display the email IDs with Edit and Confirm options"""
        edit_popup = tk.Toplevel()
        edit_popup.title("Supplier Email IDs")
        edit_popup.grab_set()

        # Entry widget for adding new email address
        tk.Label(edit_popup, text="Enter Email-Id you want to ADD: ").pack()
        email_entry = tk.Entry(edit_popup, width=40)
        email_entry.pack()

        # Add the entered email address to the Listbox
        def add_email():
            email = email_entry.get()
            if email:
                listbox.insert(tk.END, email)
                email_entry.delete(0, tk.END)

        add_button = tk.Button(edit_popup, text="Add", command=add_email)
        add_button.pack()


        # Create a Listbox to display email IDs
        tk.Label(edit_popup, text=" Email-Ids: ").pack()
        listbox = tk.Listbox(edit_popup, selectmode=tk.SINGLE, width=40, height=10)
        for email_id in email_ids:
            listbox.insert(tk.END, email_id)
        listbox.pack()

        
        # Remove selected email ID when Remove button is clicked
        def remove_email_id():
            selected_index = listbox.curselection()
            if selected_index:
                listbox.delete(selected_index)

        remove_button = tk.Button(edit_popup, text="Remove", command=remove_email_id)
        remove_button.pack()

        def send_email_with_check():
            email_list = listbox.get(0, tk.END)
            if not email_list:
                messagebox.showerror("Error", "Please add at least one email ID.")
            else:
                for email_id in email_list:
                    send_mail("Request for Quote", email_body, email_id)
                    edit_popup.destroy()
                
                messagebox.showinfo("Success", "Email sent successfully!")

        confirm_button = tk.Button(edit_popup, text="Confirm", command=send_email_with_check)
        confirm_button.pack(side=tk.RIGHT)


    def sending_email(self):
        """ This function is called as soon as the user clicks send Email in the GUI """
        if self.file_path_PR_entry.get(0):
            self.show_selection()
        else:
            messagebox.showerror("ERROR", "Upload Parts Requested File")
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)
    
    def generate_rfq(self):
        """ Main function for Generating RFQ, adding line items and creating a quote """
        if self.customer_select_box.get() and self.file_path_PR_entry.get(0):
            party_pk = self.selected_customer_partypk
            billing_details = self.data_base_conn.get_address_of_party(party_pk)
            state, country = self.data_base_conn.get_state_and_country(party_pk)
            customer_rfq_number = self.rfq_number_text.get()
            rfq_pk = self.data_base_conn.insert_into_rfq(
                party_pk, billing_details[0], billing_details[0], billing_details[1], billing_details[2], billing_details[3], billing_details[4], billing_details[5], billing_details[6], state[0][0], country[0][0], customer_rfq_number=customer_rfq_number,
            )

            user_selected_file_paths = list(self.file_path_PR_entry.get(0, tk.END) + self.file_path_PL_entry.get(0, tk.END))

            destination_paths = []
            i = 1
            y = 1
            count = 1
            my_dict = pk_info_dict(self.file_path_PR_entry.get(0, tk.END)[0])

            info_dict = part_info(self.file_path_PR_entry.get(0, tk.END)[0])

            parts = extract_from_excel(self.file_path_PR_entry.get(0, tk.END)[0], "Part")
            descriptions = extract_from_excel(self.file_path_PR_entry.get(0, tk.END)[0], "DESCRIPTION")
            part_description_data = dict(zip(parts, descriptions))

            qty = extract_from_excel(self.file_path_PR_entry.get(0, tk.END)[0], "QuantityRequired")

            qty_data = dict(zip(parts, qty))

            resricted = False

            for part_number, description in part_description_data.items():
                if self.itar_restricted_var.get():
                    destination_path = rf'y:\PDM\Restricted\{self.customer_select_box.get()}\{part_number}'
                    resricted = True
                else:
                    destination_path = rf'y:\PDM\Non-restricted\{self.customer_select_box.get()}\{part_number}'
                for file in user_selected_file_paths:
                    # folder is get or created and file is copied to this folder
                    file_path_to_add_to_rfq = self.transfer_file_to_folder(destination_path, file)
                    destination_paths.append(file_path_to_add_to_rfq)

                for file in destination_paths:
                    if count==1:
                        if resricted:
                            self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6, secure_document=1)
                        else:
                            self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6)
                count+=1

                item_pk = self.data_base_conn.get_or_create_item(part_number, description=description, purchase=0, service_item=0, manufactured_item=1)
                matching_paths = [path for path in destination_paths if part_number in path]
                 
                for url in matching_paths:
                        if resricted:
                            self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, secure_document=1)
                        else:
                            self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2)
                
                quote_pk = self.data_base_conn.create_quote(party_pk, item_pk, 0, part_number)
                self.data_base_conn.quote_operation(quote_pk)

                a = [6, 21, 22]  # IssueMat, HT, FIN
                quote_assembly_fk = []

                for x in a:
                    quote_assembly_pk = self.data_base_conn.execute_query(f"SELECT QuoteAssemblyPK FROM QuoteAssembly WHERE QuoteFK = {quote_pk} AND SequenceNumber = {x}")
            
                    quote_assembly_fk.append(quote_assembly_pk[0][0])

                rfq_line_pk = self.data_base_conn.create_rfq_line_item(item_pk, rfq_pk, i, quote_pk, quantity=qty_data[part_number])
                i+=1

                self.data_base_conn.rfq_line_quantity(rfq_line_pk, qty_data[part_number])

                if part_number in my_dict:
                    dict_values = my_dict[part_number]
                    for j, k, l in zip(dict_values, quote_assembly_fk, a):
                        if j is not None:
                            self.data_base_conn.create_bom_quote(quote_pk, j, k, l, y)
                            y+=1
                
                if part_number in info_dict:
                    dict_values = info_dict[part_number]
                    self.data_base_conn.insert_part_details_in_item(item_pk, part_number, dict_values)
                    pk_value = my_dict[part_number]
                    for j in pk_value[1:]:
                        if j:
                            self.data_base_conn.insert_part_details_in_item(j, part_number, dict_values)
                            for url in matching_paths:
                                if resricted:
                                    self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2, secure_document=1)
                                else:
                                    self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2)
                    self.data_base_conn.insert_part_details_in_item(pk_value[0], part_number, dict_values, item_type='Material')
            
            messagebox.showinfo("Success", f"RFQ generated successfully! RFQ Number: {rfq_pk}")
            answer = messagebox.askyesno("Confirmation", "Do you want to send an email to the supplier for a quote?")
            if answer:
                self.show_selection()
            
            self.file_path_PL_entry.delete(0, tk.END)
            # self.file_path_PR_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)

        else:
            messagebox.showerror("ERROR", "Select Customer/ Upload Parts Requested File")
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)


if __name__ == "__main__":
    r = RfqGen()
    r.mainloop()
