from flask import jsonify, request
from app import app
import json

@app.route('/api/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        email = data.get('email')
        password = data.get('password')

        users_data = load_users_data()
        for user in users_data:
            if user['email'] == email and user['password'] == password:
                return jsonify({'message': 'Login successful'})

        return jsonify({'message': 'Invalid credentials'}), 401
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def load_users_data():
    try:
        with open(app.config['USERS_DATA_FILE'], 'r') as file:
            users_data = json.load(file)
        return users_data
    except FileNotFoundError:
        raise Exception('Users data file not found')
    except json.JSONDecodeError:
        raise Exception('Error decoding users data JSON')
