from flask import jsonify, request
import json
import time

def load_booked_slots():
    try:
        with open('booked_slots.json', 'r') as file:
            data = json.load(file)
            return data.get('bookings', [])
    except FileNotFoundError:
        return []

bookings = load_booked_slots()

@app.route('/api/booking', methods=['POST'])
def book_slot():
    try:
        data = request.get_json()
        engineer = data.get('engineer')
        location = data.get('location')
        date = data.get('date')
        duration = data.get('duration')

        if not all([engineer, location, date, duration]):
            return jsonify({'error': 'Invalid input data'}), 400

        existing_booking = find_existing_booking(engineer, location, date, duration)
        if existing_booking:
            return jsonify({'message': 'Booking already exists', 'booking_id': existing_booking['booking_id']}), 200

        epoch_id = int(time.time())

        new_booking = {
            'booking_id': epoch_id,
            'engineer': engineer,
            'location': location,
            'date': date,
            'duration': duration
        }

        bookings.append(new_booking)
        save_booked_slots(bookings)

        return jsonify({'message': 'Booking successful', 'booking_id': epoch_id}), 201

    except Exception as e:
        return jsonify({'error': str(e)}), 500

def find_existing_booking(engineer, location, date, duration):
    for booking in bookings:
        if (
            booking['engineer'] == engineer
            and booking['location'] == location
            and booking['date'] == date
            and booking['duration'] == duration
        ):
            return booking
    return None

def save_booked_slots(bookings_data):
    with open('booked_slots.json', 'w') as file:
        json.dump({"bookings": bookings_data}, file, indent=4)


@app.route('/api/booking/fetch', methods=['POST'])
def fetch_booking():
    try:
        data = request.get_json()
        booking_id = data.get('booking_id')

        if not booking_id:
            return jsonify({'error': 'Booking ID not provided in the request body'}), 400

        booking = find_booking_by_id(booking_id)
        if booking:
            return jsonify({'message': 'Booking found', 'booking': booking}), 200
        else:
            return jsonify({'message': 'Booking not found'}), 404

    except Exception as e:
        return jsonify({'error': str(e)}), 500

def find_booking_by_id(booking_id):
    for booking in bookings:
        if booking['booking_id'] == booking_id:
            return booking
    return None

