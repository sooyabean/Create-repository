import os
import csv
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import win32api
import threading
from sqlacc_interface import SQLAccInterface

class QuoteProcessorAndInvoiceGenerator:
    def __init__(self, admin_username='ADMIN', admin_password='ADMIN'):
        """
        Initialize the quote processor
        """
        self.sqlacc = SQLAccInterface(admin_username, admin_password)
        self.QUOTES_DIR = os.path.join(os.path.dirname(__file__), 'quotes')
        os.makedirs(self.QUOTES_DIR, exist_ok=True)
        
    def create_tray_icon(self):
        """Create a tray icon for notifications"""
        def create_image():
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
                menu=(item('Quit', on_quit),)
            )
        except Exception as e:
            print(f"Error creating tray icon: {e}")

    def show_notification(self, message="Notification", title="Quote Processor"):
        """Show Windows notification"""
        try:
            win32api.MessageBox(0, message, title, 0x40 | 0x1)
        except Exception as e:
            print(f"Error showing notification: {e}")

    def run_tray_threaded(self):
        """Run tray icon in a separate thread"""
        try:
            if not hasattr(self, 'tray_icon'):
                self.create_tray_icon()
            
            tray_thread = threading.Thread(target=self.tray_icon.run)
            tray_thread.start()
        except Exception as e:
            print(f"Error running tray thread: {e}")

    def split_address_into_four(self, address):
        """Split an address into four lines"""
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

    def create_flask_app(self):
        """Create and configure Flask application"""
        app = Flask(__name__)
        CORS(app)

        @app.route("/", methods=['GET'])
        def home():
            return "Server is running! Use the /process-data endpoint to process data."

        @app.route("/process-data", methods=['POST'])
        def process_data():
            try:
                data = request.get_json()
                
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
                
                # Prepare quote data
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
                
                with open(filepath, 'w', newline='') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerow(['DATE', 'COMPANY', 'PRODUCT CODE', 'QUANTITY', 'AGENT', 
                                       'CREATEDATE', 'ADDRESS', 'POSTCODE', 'CITY', 'STATE', 
                                       'ATTENTION', 'MOBILE', 'PHONE1', 'EMAIL', 'FAX1'])
                    csv_writer.writerows(quote_data)

                # Process CSV and generate invoices
                processing_result = self.process_csv_to_invoices(filepath)

                self.show_notification(
                    f"Successfully added new quote for: {company_name}", 
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

    def process_csv_to_invoices(self, csv_file_path):
        """Process CSV file and create invoices"""
        processing_results = {
            'total_records': 0,
            'processed_records': 0,
            'skipped_records': 0
        }

        try:
            with open(csv_file_path, 'r') as csvfile:
                csv_reader = csv.reader(csvfile)
                next(csv_reader)  # Skip header

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
                    }# Find company code using SQLAcc interface
                    company_code = self.sqlacc.find_company_code(
                        company_details['companyName'], 
                        company_details['agent'], 
                        company_details
                    )
                    
                    if isinstance(company_code, dict) and company_code.get('status') == 'typo':
                        # Handle possible typos
                        self.show_notification(
                            company_code['message'],
                            "Company Name Typo Detected"
                        )
                        processing_results['skipped_records'] += 1
                        continue

                    if company_code:
                        # Add quotation using SQLAcc interface
                        if self.sqlacc.add_quotation(record, company_code):
                            processing_results['processed_records'] += 1
                        else:
                            processing_results['skipped_records'] += 1
                            print(f"Failed to add quotation for company: {record[1]}")
                    else:
                        processing_results['skipped_records'] += 1
                        print(f"Skipping record due to missing company code: {record[1]}")

        except Exception as e:
            print(f"Error processing CSV: {e}")
            processing_results['error'] = str(e)

        return processing_results

    def run_server(self, host='0.0.0.0', port=3000, debug=True):
        """
        Run the Flask server
        
        Args:
            host (str): Host to bind the server
            port (int): Port to run the server on
            debug (bool): Enable debug mode
        """
        app = self.create_flask_app()
        # Start the tray icon in a separate thread
        self.run_tray_threaded()
        # Run the Flask application
        app.run(host=host, port=port, debug=debug)

def main():
    """
    Main entry point of the application
    """
    try:
        processor = QuoteProcessorAndInvoiceGenerator()
        processor.run_server()
    except Exception as e:
        print(f"Application startup error: {e}")

if __name__ == "__main__":
    main()
