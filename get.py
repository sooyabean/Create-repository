from flask import Flask, jsonify, request
import requests
from typing import List, Dict
import time
import json

app = Flask(__name__)
URL = 'http://192.168.0.32:5001/receive-data'

# Store received data in a list
received_data: List[Dict] = []

@app.route('/')
def testing():
    return 'Webhook Server Running'

def share_data_to_server(data: Dict) -> tuple:
    """
    Share data to external server and return response status and message
    """
    try:
        # Print the data being sent for debugging
        print("Attempting to send data:", json.dumps(data, indent=2))
        
        response = requests.post(
            URL,
            json=data,
            headers={'Content-Type': 'application/json'},
            timeout=10  # Add timeout
        )
        
        # Print response for debugging
        print(f"Response status: {response.status_code}")
        print(f"Response content: {response.text}")
        
        if response.status_code == 200:
            return True, "Data forwarded successfully"
        return False, f"Failed to share data. Status code: {response.status_code}, Response: {response.text}"
        
    except requests.exceptions.ConnectionError:
        return False, f"Failed to connect to receiving server at {URL}. Is it running?"
    except requests.exceptions.Timeout:
        return False, "Request timed out"
    except Exception as e:
        return False, f"Error sharing data: {str(e)}"

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        # Print raw request data for debugging
        print("Received request headers:", dict(request.headers))
        print("Received raw data:", request.get_data(as_text=True))
        
        # Get JSON data from the request
        data = request.get_json(force=True)  # Use force=True to handle string-encoded JSON
        
        if not data:
            return jsonify({
                "status": "error",
                "message": "No data received"
            }), 400
        
        # Print parsed data
        print("Parsed data:", json.dumps(data, indent=2))
        
        # Add timestamp to the data
        data['received_at'] = time.time()
        
        # Store the data in our list
        received_data.append(data)
        
        # Share data to server
        success, message = share_data_to_server(data)
        
        if success:
            return jsonify({
                "status": "success",
                "message": message,
                "data_position": len(received_data) - 1
            }), 200
        else:
            return jsonify({
                "status": "partial_success",
                "message": f"Data stored but sharing failed: {message}",
                "data_position": len(received_data) - 1
            }), 200  # Return 200 even for partial success
            
    except json.JSONDecodeError as e:
        return jsonify({
            "status": "error",
            "message": f"Invalid JSON format: {str(e)}"
        }), 400
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"Internal error: {str(e)}"
        }), 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)
