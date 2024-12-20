from flask import Flask, jsonify, request
import requests
from typing import List, Dict
import time

app = Flask(__name__)
URL = 'http://localhost:5001/receive-data'

# Store received data in a list
received_data: List[Dict] = []

@app.route('/')
def testing():
    return 'Hiiiii'

def share_data_to_server(data: Dict) -> tuple:
    """
    Share data to external server and return response status and message
    """
    try:
        response = requests.post(
            URL,
            json=data,
            headers={'Content-Type': 'application/json'}
        )
        
        if response.status_code == 200:
            return True, "Data forwarded successfully"
        return False, f"Failed to share data. Status code: {response.status_code}"
        
    except requests.exceptions.RequestException as e:
        return False, f"Failed to connect to receiving service: {str(e)}"
    except Exception as e:
        return False, f"Error sharing data: {str(e)}"

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        # Get JSON data from the request
        data = request.json
        
        if not data:
            return jsonify({
                "status": "error",
                "message": "No data received"
            }), 400
            
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
                "data_position": len(received_data) - 1  # Return position in the list
            }), 200
        else:
            return jsonify({
                "status": "partial_success",
                "message": f"Data stored but sharing failed: {message}",
                "data_position": len(received_data) - 1
            }), 500
            
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"Internal error: {str(e)}"
        }), 500

@app.route('/stored-data', methods=['GET'])
def get_stored_data():
    """Endpoint to view all stored data"""
    return jsonify({
        "status": "success",
        "count": len(received_data),
        "data": received_data
    })

@app.route('/retry-share/<int:position>', methods=['POST'])
def retry_share_data(position: int):
    """Endpoint to retry sharing specific data"""
    try:
        if position < 0 or position >= len(received_data):
            return jsonify({
                "status": "error",
                "message": "Invalid position"
            }), 400
            
        data = received_data[position]
        success, message = share_data_to_server(data)
        
        if success:
            return jsonify({
                "status": "success",
                "message": message
            }), 200
        else:
            return jsonify({
                "status": "error",
                "message": message
            }), 500
            
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"Internal error: {str(e)}"
        }), 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)
