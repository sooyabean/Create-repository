from flask import Flask, request, jsonify
from flask_cors import CORS

class GetDataFromZoho:
    def create_flask_app(self):
        app = Flask(__name__)
        CORS(app)  # Enable CORS for cross-origin requests

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

        return app

def main():
    processor = GetDataFromZoho()
    app = processor.create_flask_app()
    # Run the app on all available network interfaces
    app.run(debug=True, host='0.0.0.0', port=3000)

if __name__ == "__main__":
    main()
