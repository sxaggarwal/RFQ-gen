# all tkinter modules that go on the screen
from src.Mie_trak_connection import MieTrak
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from src.pull_excel_data import extract_from_excel
import os, shutil
from src.part_info_dict import pk_info_dict
from src.docparser import boeing_pdf_converter
from src.seperating_file import get_pl_path
from src.data_cleaner_boeing import capture_data
from src.extract_data_from_pl import extract_finish_codes_from_file, extract_dash_number

class RfqGen(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title = "RFQGen"
        self.geometry("600x400")

        self.data_base_conn = MieTrak()

        ## assign - sid : add display for customer, get pk, pull data from SQL to display in box. (start box from (2, 1)). 
        ## task 2 - reorder the widgets in a way that make the most sense. 
        self.customer_names = [customer[0] for customer in self.data_base_conn.execute_query("select name from party")]
        query = "select partypk, name from party"
        self.data = self.data_base_conn.execute_query(query)
        self.customer_names = [d[1] for d in self.data]
        self.customer_number_to_partypk = {i: d[0] for i, d in enumerate(self.data, start=1)}

        self.make_combobox()

    def make_combobox(self):
        # Customer select combobox
        tk.Label(self, text="Select Customer: ").grid(row=0, column=0)
        self.customer_select_box = ttk.Combobox(self, values=self.customer_names, state = "readonly")
        self.customer_select_box.grid(row=1, column=0)
        
        self.customer_info_text = tk.Text(self, height=4, width = 30)
        self.customer_info_text.grid(row=1, column=1)

        # Bind the combobox selection event to update customer information
        self.customer_select_box.bind("<<ComboboxSelected>>", self.update_customer_info)

        self.output_text = tk.Text(self, height=4, width=30)
        self.output_text.grid(row=7, column=1)

        # Entrybox for the requested parts in an excel file. (upload for excel file) 
        tk.Label(self, text="Parts Requested File").grid(row=2, column=0)
        self.file_path_PR_entry = tk.Listbox(self, height=2, width = 50)
        self.file_path_PR_entry.grid(row=3, column=0)

        # this should be at the bottom
        browse_button_1 = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("Excel files", self.file_path_PR_entry))
        browse_button_1.grid(row=4, column=0)

        #Selection/ Upload for PartList
        tk.Label(self, text="Part Lists File (PL)").grid(row=5, column=0)
        self.file_path_PL_entry = tk.Listbox(self, height=2, width = 50)
        self.file_path_PL_entry.grid(row=6, column=0)

        browse_button_part_list = tk.Button(self, text="Browse Files", command= lambda: self.browse_files_parts_requested("All files", self.file_path_PL_entry))
        browse_button_part_list.grid(row=7, column=0)

        # main button
        generate_button = tk.Button(self, text="Generate RFQ", command=self.generate_rfq)
        generate_button.grid(row=8, column=0)
    
    def update_customer_info(self, event= None):
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
        selected_customer_number = selected_customer_index + 1  # Adding 1 because indices start from 0
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

    def generate_rfq(self):
        """ Main function for Generating RFQ, adding line items and creating a quote """
        if self.customer_select_box.get() and self.file_path_PR_entry.get(0):
            self.output_text.delete(1.0, tk.END)
            party_pk = self.selected_customer_partypk
            billing_details = self.data_base_conn.get_address_of_party(party_pk)
            state, country = self.data_base_conn.get_state_and_country(party_pk)
            
            self.data_base_conn.insert_into_rfq(
                party_pk, billing_details[0], billing_details[0], billing_details[1], billing_details[2], billing_details[3], billing_details[4], billing_details[5], billing_details[6], state[0][0], country[0][0]
            )

            rfq_pk = self.data_base_conn.get_rfq_pk()

            user_selected_file_paths = list(self.file_path_PR_entry.get(0, tk.END) + self.file_path_PL_entry.get(0, tk.END))

            destination_paths = []
            i = 1
            y = 1
            # my_dict = {'N536T5506-204-00': [102990, 102999, None], 'N533T5501-200-00': [102961, 102996, None], 'N533T5501-201-00': [102962, None, None], 'N533T5501-202-00': [102963, None, None], 'N533T5501-204-00': [102964, None, None], 'N533T5501-205-00': [102965, None, None], 'N533T5501-206-00': [102966, None, None], 'N533T5501-207-00': [102967, None, None], 'N533T5503-200-00': [102968, None, None], 'N533T5503-202-00': [102969, None, None], 'N533T5503-204-00': [102970, None, None], 'N533T5503-206-00': [102971, None, None], 'N533T5506-202-00': [102972, None, None], 'N533T5506-203-00': [102973, 102997, None], 'N536T5501-200-00': [102974, 102998, None], 'N536T5501-202-00': [102975, None, 103019], 'N536T5501-204-00': [102976, None, 103020], 'N536T5501-206-00': [102977, None, 103021], 'N536T5501-208-00': [102978, None, 103022], 'N536T5501-210-00': [102979, None, 103023], 'N536T5501-212-00': [102980, None, 103024], 'N536T5501-213-00': [102981, None, 103025], 'N536T5501-214-00': [102982, None, 103026], 'N536T5501-215-00': [102983, None, 103027], 'N536T5503-200-00': [102984, None, 103028], 'N536T5503-202-00': [102985, None, 103029], 'N536T5503-204-00': [102986, None, 103030], 'N536T5503-206-00': [102987, None, 103031], 'N536T5503-208-00': [102988, None, 103032], 'N536T5506-202-00': [102989, None, 103033]}
            my_dict = pk_info_dict(self.file_path_PR_entry.get(0, tk.END)[0])
            print(my_dict)

            parts = extract_from_excel(self.file_path_PR_entry.get(0, tk.END)[0], "Part")
            # parts = [value for value in parts_rough if value is not nan] #There are some values as 'nan' because of the excel formatting so removing that.
            descriptions = extract_from_excel(self.file_path_PR_entry.get(0, tk.END)[0], "DESCRIPTION")
            # descriptions = [value for value in descriptions_rough if value is not nan]
            part_description_data = dict(zip(parts, descriptions))

            for part_number, description in part_description_data.items():
                destination_path = rf'y:\PDM\Non-restricted\{self.customer_select_box.get()}\{part_number}'
                for file in user_selected_file_paths:
                    # folder is get or created and file is copied to this folder
                    file_path_to_add_to_rfq = self.transfer_file_to_folder(destination_path, file)
                    # we will then add the filepath to the rfq created above
                    destination_paths.append(file_path_to_add_to_rfq)
                    self.data_base_conn.upload_documents(file_path_to_add_to_rfq, rfq_fk=rfq_pk)

                item_pk = self.data_base_conn.get_or_create_item(part_number, description= description)
                matching_paths = [path for path in destination_paths if part_number in path]
                # print(matching_paths)

                 
                for url in matching_paths:
                    self.data_base_conn.upload_documents(url, item_fk=item_pk)
                
                quote_pk = self.data_base_conn.create_quote(party_pk, item_pk, 0, part_number)
                self.data_base_conn.quote_operation(quote_pk)

                a = [6,21,22] #  IssueMat, HT, FIN
                quote_assembly_fk = []
                # TODO: get the quote assembly pk for finish and heat treat and material issue. (22,21 and 6) [match with sequence number and quotefk = quotepk]
                for x in a:
                    quote_assembly_pk = self.data_base_conn.execute_query(f"SELECT QuoteAssemblyPK FROM QuoteAssembly WHERE QuoteFK = {quote_pk} AND SequenceNumber = {x}")
                    # print(quote_assembly_pk)
                    quote_assembly_fk.append(quote_assembly_pk[0][0])
                
                self.data_base_conn.create_rfq_line_item(item_pk, rfq_pk, i, quote_pk)
                i+=1

                if part_number in my_dict:
                    dict_values = my_dict[part_number]
                    for j,k,l in zip(dict_values, quote_assembly_fk, a):
                        # print(j,k,l)
                        if j != None:
                            self.data_base_conn.create_bom_quote(quote_pk, j, k, l, y)
                            y+=1

                # TODO: Converting and cleanning a pdf.
                # paths_with_pl = get_pl_path(destination_paths)
                # for path in paths_with_pl:
                #     if part_number in path:
                #         boeing_pdf_converter(path, f"{part_number}_raw_PL.txt")
                #         capture_data(f"{part_number}_raw_PL.txt", f"{part_number}_PL.txt")
                
                # # TODO: FEATURE NOT COMPLETE, COMMENTED TO PUSH TO GIT
                # dash_number = extract_dash_number(part_number)
                # result, part_info_list = extract_finish_codes_from_file(f"{part_number}_PL.txt", dash_number)
                # if result:
                #     output_message = f"Finish codes for part {part_number}: {result}\n"
                #     if part_info_list:
                #         output_message += "Material Information:\n"
                #         for part_info in part_info_list:
                #             output_message += f"{part_info}\n"
                #     else:
                #         output_message += "No material information available for this part.\n"
                # else:
                #     output_message = f"No Finish codes for {part_number}\n"
                
                # self.output_text.insert(tk.END, output_message)

                # creating router
                # router_pk = self.data_base_conn.create_router(party_pk, item_pk)
                # print(router_pk)
                # self.data_base_conn.router_work_center(router_pk)
        
            
            messagebox.showinfo("Success", f"RFQ generated successfully! RFQ Number: {rfq_pk}")
            # self.customer_select_box.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)

        else:
            messagebox.showerror("ERROR", "Select Customer/ Upload Parts Requested File")
            self.file_path_PR_entry.delete(0, tk.END)
            # self.customer_select_box.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)


if __name__ == "__main__":
    r = RfqGen()
    r.mainloop()
