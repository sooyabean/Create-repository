import os
import csv
import re
import traceback
import win32com.client
import pythoncom
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime
from difflib import get_close_matches
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import win32api
import threading
import time

class QuoteProcessorAndInvoiceGenerator:
    def __init__(self, admin_username='ADMIN', admin_password='ADMIN'):
        """
        Initialize the quote processor with SQLAcc credentials and file management
        """
        # SQLAcc connection properties
        self.com_server = None
        self.admin_username = admin_username
        self.admin_password = admin_password

        # Quote storage configuration
        self.QUOTES_DIR = os.path.join(os.path.dirname(__file__), 'quotes')
        os.makedirs(self.QUOTES_DIR, exist_ok=True)
        
    def create_tray_icon(self):
        """
        Create a tray icon for notifications
        """
        def create_image():
            # Create an image for the tray icon
            width = 64
            height = 64
            image = Image.new('RGB', (width, height), (255, 255, 255))
            draw = ImageDraw.Draw(image)
            draw.rectangle((0, 0, width, height), fill=(0, 128, 0))
            return image

        def on_quit(icon, item):
            icon.stop()

        try:
            self.tray_icon = pystray.Icon(
                "QuoteProcessor", 
                create_image(), 
                menu=(
                    item('Quit', on_quit),
                )
            )
        except Exception as e:
            self.logger.error(f"Error creating tray icon: {e}")

    def show_notification(self, message="Notification", title="Quote Processor"):
        """
        Show Windows notification
        
        Args:
            message (str): Notification message
            title (str): Notification title
        """
        try:
            win32api.MessageBox(0, message, title, 0x40 | 0x1)
        except Exception as e:
            self.logger.error(f"Error showing notification: {e}")

    def run_tray_threaded(self):
        """
        Run tray icon in a separate thread
        """
        try:
            if not self.tray_icon:
                self.create_tray_icon()
            
            tray_thread = threading.Thread(target=self.tray_icon.run)
            tray_thread.start()
        except Exception as e:
            self.logger.error(f"Error running tray thread: {e}")

    def create_sqlacc_server(self):
        """
        Create and login to SQLAcc business application
        
        Returns:
            bool: True if server creation and login is successful, False otherwise
        """
        try:
            pythoncom.CoInitialize()
            # Create COM object for SQLAcc
            self.com_server = win32com.client.Dispatch("SQLAcc.BizApp")
            
            if not self.com_server:
                print("Failed to create SQLAcc.BizApp object")
                return False

            # Login if not already logged in
            if not self.com_server.IsLogin:
                try:
                    self.com_server.Login(self.admin_username, self.admin_password)
                except Exception as login_error:
                    print(f"Login Error: {login_error}")
                    return False
            
            return True
        except Exception as e:
            print(f"ActiveX Creation Error: {e}")
            print(f"Error Trace: {traceback.format_exc()}")
            return False

    def create_flask_app(self):
        """
        Create and configure Flask application for quote processing
        
        Returns:
            Flask: Configured Flask application
        """
        app = Flask(__name__)
        CORS(app)

        @app.route("/", methods=['GET'])
        def home():
            """Basic server health check endpoint"""
            return "Server is running! Use the /process-data endpoint to process data."

        @app.route("/process-data", methods=['POST'])
        def process_data():
            """
            Process quote data, save to CSV, and generate SQLAcc invoice
            """
            try:
                data = request.get_json()
                print(data)
                
                quote_date = data.get('quoteDate')
                company_name = data.get('companyName')
                items = data.get('items')
                agent = data.get('agent')
                company_details = data.get('companyDetails')
                
                if not all([quote_date, company_name, items, agent, company_details]):
                    return jsonify({
                        'status': 'error',
                        'message': 'Missing required fields'
                    }), 400
                
                # Prepare quote data for CSV
                quote_data = []
                for item in items:
                    quote_data.append([
                        quote_date,
                        company_name,
                        item['productCode'],
                        item['quantity'],
                        agent,
                        company_details['createDate'],
                        company_details['address'],
                        company_details['postcode'],
                        company_details['city'],
                        company_details['state'],
                        company_details['attention'],
                        company_details['mobile'],
                        company_details['phone1'],
                        company_details['email'],
                        company_details['fax1']
                    ])
                
                # Generate CSV file
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f'quote_{timestamp}.csv'
                filepath = os.path.join(self.QUOTES_DIR, filename)
                
                # Write to CSV
                with open(filepath, 'w', newline='') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    # Write header         0         1             2             3         4          5            6           7        8        9          10          11       12        13      14
                    csv_writer.writerow(['DATE', 'COMPANY', 'PRODUCT CODE', 'QUANTITY', 'AGENT', 'CREATEDATE', 'ADDRESS', 'POSTCODE', 'CITY', 'STATE', 'ATTENTION', 'MOBILE', 'PHONE1', 'EMAIL', 'FAX1'])
                    # Write data
                    csv_writer.writerows(quote_data)
                
                print(f"CSV file has been written to {filepath}")

                # Process the CSV and generate SQLAcc invoices
                processing_result = self.process_csv_to_invoices(filepath)

                self.show_notification(
                    f"Sucessfully added new quote for: {company_name}", 
                    "Success"
                )
                return jsonify({
                    'status': 'success', 
                    'message': 'Quote data saved to CSV and processed in SQLAcc',
                    'filename': filename,
                    'sqlacc_processing': processing_result
                }), 200
            
            except Exception as error:
                print(f"Error processing data: {error}")
                return jsonify({
                    'status': 'error', 
                    'message': str(error)
                }), 500

        return app
    
    def normalize_company_name(self, name):
        """
        Enhanced normalization with special character tracking
        """
        # Log original name
        original_name = name
        
        # Remove special characters
        normalized_name = re.sub(r'[^a-zA-Z0-9\s]', '', name.strip().lower())
        
        return normalized_name

    def split_address_into_four(self, address):
        """
        Split an address with specific length handling for each part
    
        Args:
            address (str): Full address string
    
        Returns:
            list: A list of 4 address lines
        """
        address_parts = [part.strip() for part in address.split(',')]

        def get_address_part(parts, max_chars=60):
            if not parts:
                return '', []
        
            if len(parts[0]) > max_chars:
                return parts[0][:max_chars], parts[1:]

            if len(parts) > 1 and len(parts[0]) + len(parts[1] + ', ') > max_chars:
                return parts[0], parts[1:]
        
            if len(parts) > 1:
                return f"{parts[0]}, {parts[1]}", parts[2:]
        
            return parts[0], parts[1:]

        result = ['', '', '', '']
    
        result[0], remaining = get_address_part(address_parts)
    
        if remaining:
            result[1], remaining = get_address_part(remaining)
    
        if remaining:
            result[2], remaining = get_address_part(remaining)
    
        if remaining:
            result[3] = ', '.join(remaining)
    
        return result

    def find_company_code(self, company_name, agent, company_details):
        """
        Find company code based on company name
        
        Args:
            company_name (str): Name of the company to search for
            agent (str): The agent associated with the company
            company_details (dict): Dictionary containing detailed information about the company

        Returns:
            str, dict, or None: 
                - Company code as a string if found.
                - A dictionary with a status and message suggesting possible typos.
                - None if no match is found or creation fails.
        """
        try:
            # Initialize the SQLAcc server
            if not self.create_sqlacc_server():
                print("Failed to connect to SQLAcc")
                return None
            
            biz_object = self.com_server.BizObjects.Find('AR_CUSTOMER')

            # Normalize the input company name for comparison
            normalized_input_name = self.normalize_company_name(company_name)

            # Retrieve all customer records from SQLAcc
            result = biz_object.Select(
                "CODE, COMPANYNAME", 
                "", 
                "", 
                "SD", 
                ",", 
                ""
            )
            
            if not result:
                print("No records returned from SQLAcc")
                return None

            company_data = []
            # Parse the results into a list of tuples
            for line in result.split("\n")[1:]:  # Skip header row
                if line:
                    code, db_name = line.split(",", 1)  # Split into code and company name
                    normalized_db_name = self.normalize_company_name(db_name)
                    company_data.append((code, normalized_db_name, db_name))

            # Attempt exact match
            for code, normalized_db_name, db_name in company_data:
                if normalized_input_name == normalized_db_name:
                    print(f"Exact match found for company: {db_name}, Code:{code}")
                    return code

            # Attempt fuzzy matching for typos
            possible_matches = get_close_matches(normalized_input_name, [item[1] for item in company_data], n=3, cutoff=0.8)
            if possible_matches:
                similar_names = [item[2] for item in company_data if item[1] in possible_matches]
                
                print(f"Typo Detected! Possible matches: {', '.join(similar_names)}")
                self.show_notification(
                    f"Possible company matches:\n{', '.join(similar_names)}", 
                    "Typo Detected"
                )
                
                return {
                    "status": "typo",
                    "message": f"Did you mean one of these? {', '.join(similar_names)}"
                }

            # If no match, prepare to create a new customer
            print(f"No matching company found. Preparing to create new customer: {company_name}")
        
            # Create new customer in SQLAcc
            self.create_sqlacc_server()
            main_dataset = biz_object.DataSets.Find('MainDataSet')
            branch_dataset = biz_object.DataSets.Find('cdsBranch')
        
            biz_object.New()
            main_dataset.FindField('CompanyName').value = company_details['companyName']
            main_dataset.FindField('Agent').value = company_details['agent']
            main_dataset.FindField('CreationDate').value = company_details['createDate']
        
            branch_dataset.Edit()
            branch_dataset.FindField('DtlKey').value = -1
            branch_dataset.FindField('Address1').value = company_details['address1']
            branch_dataset.FindField('Address2').value = company_details['address2']
            branch_dataset.FindField('Address3').value = company_details['address3']
            branch_dataset.FindField('Address4').value = company_details['address4']
            branch_dataset.FindField('Postcode').value = company_details['postcode']
            branch_dataset.FindField('City').value = company_details['city']
            branch_dataset.FindField('State').value = company_details['state']
            branch_dataset.FindField('Attention').value = company_details['attention']
            branch_dataset.FindField('Phone1').value = company_details['phone1']
            branch_dataset.FindField('Mobile').value = company_details['mobile']
            branch_dataset.FindField('Fax1').value = company_details['fax1']
            branch_dataset.FindField('Email').value = company_details['email']
        
            biz_object.Save()
            print("Customer saved")
        
            # Retrieve the newly created company code
            escaped_company_name = company_details['companyName'].replace("'", "''")  # Escape single quotes
            result = biz_object.Select(
                "CODE",
                f"COMPANYNAME = '{escaped_company_name}'",
                "",
                "SD",
                ",",
                ""
            )
        
            if result:
                lines = result.split("\n")
                if len(lines) > 1 and lines[1]:
                    print(f"New company code found: {lines[1]}")
                    return lines[1]
        
            print("Error retrieving the code for the newly created company")
            self.show_notification(
                "Error retrieving new company code", 
                "Company Code Error"
            )
            return None
    
        except Exception as e:
            print(f"Error finding company code: {e}")
            print(f"Error Trace: {traceback.format_exc()}")
            return None
        
    def add_quotation(self, record, company_code):
        """
        Add a quotation to SQLAcc
        
        Args:
            record (list): CSV record details
            company_code (str): Company code
            last_docno (str, optional): Last document number
        
        Returns:
            bool: True if quotation added successfully, False otherwise
        """
        try:
            self.create_sqlacc_server()
            doc_date, company_name, item_code, qty = record[:4]
            
            biz_object = self.com_server.BizObjects.Find('SL_QT')
            main_dataset = biz_object.DataSets.Find('MainDataSet')
            
            biz_object.New()
            main_dataset.FindField('DocDate').value = doc_date
            main_dataset.FindField('Code').value = company_code.strip()
            main_dataset.FindField('Description').value = "Quotation"

            detail_dataset = biz_object.DataSets.Find('cdsDocDetail')
            detail_dataset.Append()
            detail_dataset.FindField('DtlKey').value = -1
            detail_dataset.FindField('DocKey').value = -1
            detail_dataset.FindField('ItemCode').value = item_code
            detail_dataset.FindField('Qty').value = qty
            detail_dataset.Post()

            biz_object.Save()
            print("Quotation has been saved successfully.")
            return True
        

        except Exception as e:
            print(f"Error in AddQuotation: {e}")
            return False

    def process_csv_to_invoices(self, csv_file_path):
        """
        Process entire CSV file and create invoices
        
        Args:
            csv_file_path (str): Path to CSV file
        
        Returns:
            dict: Processing result summary
        """
        processing_results = {
            'total_records': 0,
            'processed_records': 0,
            'skipped_records': 0
        }

        try:
            # Read CSV file
            with open(csv_file_path, 'r') as csvfile:
                csv_reader = csv.reader(csvfile)
                # Skip header
                next(csv_reader)

                # Process each record
                for record in csv_reader:
                    processing_results['total_records'] += 1
                    
                    address_lines = self.split_address_into_four(record[6])
                    
                    company_details = {
                        'companyName': record[1],
                        'agent': record[4],
                        'createDate': record[5],
                        'address1': address_lines[0],
                        'address2': address_lines[1],
                        'address3': address_lines[2],
                        'address4': address_lines[3],
                        'postcode': record[7],
                        'city': record[8],
                        'state': record[9],
                        'attention': record[10],
                        'mobile': record[11],
                        'phone1': record[12],
                        'email': record[13],
                        'fax1': record[14]
                    }
                    
                    # Find company code
                    company_code = self.find_company_code(company_details['companyName'], company_details['agent'], company_details)
                    
                    if company_code:
                        # Add quotation
                        if self.add_quotation(record, company_code):
                            processing_results['processed_records'] += 1
                        else:
                            processing_results['skipped_records'] += 1
                    else:
                        processing_results['skipped_records'] += 1
                        print(f"Skipping record due to missing company code: {record[1]}")

        except Exception as e:
            print(f"Error processing CSV: {e}")
            processing_results['error'] = str(e)

        return processing_results
    
    def split_address_into_four(self, address):
        """
        Split an address with specific length handling for each part
    
        Args:
            address (str): Full address string
    
        Returns:
            list: A list of 4 address lines
        """
        address_parts = [part.strip() for part in address.split(',')]

        def get_address_part(parts, max_chars=60):
            if not parts:
                return '', []
        
            if len(parts[0]) > max_chars:
                return parts[0][:max_chars], parts[1:]

            if len(parts) > 1 and len(parts[0]) + len(parts[1] + ', ') > max_chars:
                return parts[0], parts[1:]
        
            if len(parts) > 1:
                return f"{parts[0]}, {parts[1]}", parts[2:]
        
            return parts[0], parts[1:]

        result = ['', '', '', '']
    
        result[0], remaining = get_address_part(address_parts)
    
        if remaining:
            result[1], remaining = get_address_part(remaining)
    
        if remaining:
            result[2], remaining = get_address_part(remaining)
    
        if remaining:
            result[3] = ', '.join(remaining)
    
        return result

    def run_server(self, host='0.0.0.0', port=3000, debug=True):
        """
        Run the Flask server
        
        Args:
            host (str): Host to bind the server
            port (int): Port to run the server on
            debug (bool): Enable debug mode
        """
        app = self.create_flask_app()
        app.run(host=host, port=port, debug=debug)

def main():
    processor = QuoteProcessorAndInvoiceGenerator()
    processor.run_server()

if __name__ == "__main__":
    main()
    
# boleh guna tapi kena buat dia boleh store new customer data punya dulu. DONE
# once dah boleh baru boleh usha how to store one with lots of items, etc. 
# lepas dah siap ni weekend tengok pasal dynamo
# kalau free or rajin tengok pasal acc autodesk and power bi 
# oh yeee cer update pasal project buat boleh tengok through the yearrrrr

# buat chapter 2 and correction chapter 1 !!!!!! DONE
# terus buat chapter 3 mana yang siap dulu
# minggu ni fokus on quotation + project !!!!!
# weekend buat kerja kat luar (lalaport)
# make sure kerja siap before date