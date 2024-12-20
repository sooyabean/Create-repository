from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)  # Create the app instance at module level
CORS(app)

@app.route("/", methods=['GET'])
def home():
    return "Server is running"

@app.route("/webhook", methods=['POST'])
def process_data():
    try:
        data = request.get_json()
        print(f"Received data: {data}")  # Log the received data
        # Process the data here (optional)
        return jsonify({"message": "Data received successfully", "data": data}), 200
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": "Failed to process data"}), 400

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=3000)
