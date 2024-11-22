from flask import Blueprint, jsonify, request
from datetime import datetime, timedelta
from sqlalchemy import cast, Date, func, and_, or_, extract
from model import db, RoomReservation, VenueReservation, Room, RoomType, Receipt, GuestDetails

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route('/api/dashboardData', methods=['GET'])
def get_dashboard_data():
    try:
        end_date_str = request.args.get('endDate')
        start_date_str = request.args.get('startDate')
        view_mode = request.args.get('viewMode', 'daily')

        # Ensure proper date parsing
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date() if end_date_str else datetime.now().date()
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date() if start_date_str else (end_date - timedelta(days=7))

        print(f"Start Date: {start_date}, End Date: {end_date}, View Mode: {view_mode}")

        # Generate date range for occupancy data
        date_range = []
        current_date = start_date
        while current_date <= end_date:
            date_range.append(current_date)
            current_date += timedelta(days=1)

        # Query reservations within the date range
        room_reservations = RoomReservation.query.filter(
            or_(
                and_(
                    cast(RoomReservation.room_reservation_booking_date_start, Date) <= end_date,
                    cast(RoomReservation.room_reservation_booking_date_start, Date) >= start_date
                ),
                and_(
                    cast(RoomReservation.room_reservation_booking_date_end, Date) >= start_date,
                    cast(RoomReservation.room_reservation_booking_date_end, Date) <= end_date
                )
            )
        ).all()

        venue_reservations = VenueReservation.query.filter(
            or_(
                and_(
                    cast(VenueReservation.venue_reservation_booking_date_start, Date) <= end_date,
                    cast(VenueReservation.venue_reservation_booking_date_start, Date) >= start_date
                ),
                and_(
                    cast(VenueReservation.venue_reservation_booking_date_end, Date) >= start_date,
                    cast(VenueReservation.venue_reservation_booking_date_end, Date) <= end_date
                )
            )
        ).all()

        # Calculate previous period
        period_length = (end_date - start_date).days + 1
        prev_end_date = start_date - timedelta(days=1)
        prev_start_date = prev_end_date - timedelta(days=period_length - 1)

        # Previous period queries
        prev_room_reservations = RoomReservation.query.filter(
            or_(
                and_(
                    cast(RoomReservation.room_reservation_booking_date_start, Date) <= prev_end_date,
                    cast(RoomReservation.room_reservation_booking_date_start, Date) >= prev_start_date
                ),
                and_(
                    cast(RoomReservation.room_reservation_booking_date_end, Date) >= prev_start_date,
                    cast(RoomReservation.room_reservation_booking_date_end, Date) <= prev_end_date
                )
            )
        ).all()

        prev_venue_reservations = VenueReservation.query.filter(
            or_(
                and_(
                    cast(VenueReservation.venue_reservation_booking_date_start, Date) <= prev_end_date,
                    cast(VenueReservation.venue_reservation_booking_date_start, Date) >= prev_start_date
                ),
                and_(
                    cast(VenueReservation.venue_reservation_booking_date_end, Date) >= prev_start_date,
                    cast(VenueReservation.venue_reservation_booking_date_end, Date) <= prev_end_date
                )
            )
        ).all()

        # Calculate metrics
        def calculate_percentage_change(current, previous):
            if previous == 0:
                return 100 if current > 0 else 0
            return ((current - previous) / previous) * 100

        # Current period calculations
        current_bookings = len(room_reservations) + len(venue_reservations)
        current_revenue = db.session.query(func.sum(Receipt.receipt_total_amount)).filter(
            cast(Receipt.receipt_date, Date) >= start_date,
            cast(Receipt.receipt_date, Date) <= end_date
        ).scalar() or 0

        current_guests = db.session.query(func.count(func.distinct(GuestDetails.guest_id))).filter(
            or_(
                GuestDetails.guest_id.in_([r.guest_id for r in room_reservations]),
                GuestDetails.guest_id.in_([v.guest_id for v in venue_reservations])
            )
        ).scalar()

        # Previous period calculations
        prev_bookings = len(prev_room_reservations) + len(prev_venue_reservations)
        prev_revenue = db.session.query(func.sum(Receipt.receipt_total_amount)).filter(
            cast(Receipt.receipt_date, Date) >= prev_start_date,
            cast(Receipt.receipt_date, Date) <= prev_end_date
        ).scalar() or 0

        prev_guests = db.session.query(func.count(func.distinct(GuestDetails.guest_id))).filter(
            or_(
                GuestDetails.guest_id.in_([r.guest_id for r in prev_room_reservations]),
                GuestDetails.guest_id.in_([v.guest_id for v in prev_venue_reservations])
            )
        ).scalar()

        # Available rooms calculations
        total_rooms = db.session.query(func.count(Room.room_id)).scalar() or 0
        
        # Occupancy data
        if view_mode == 'monthly':
            occupancy_data = []
            current_date = start_date.replace(day=1)
            while current_date <= end_date:
                month_end = (current_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
                month_end = min(month_end, end_date)
                
                occupied_count = db.session.query(func.count(RoomReservation.room_reservation_id)).filter(
                    cast(RoomReservation.room_reservation_booking_date_start, Date) <= month_end,
                    cast(RoomReservation.room_reservation_booking_date_end, Date) >= current_date
                ).scalar() or 0
                
                occupancy_data.append({
                    "date": current_date.strftime('%Y-%m'),
                    "occupancy": occupied_count
                })
                
                current_date = (current_date.replace(day=28) + timedelta(days=4)).replace(day=1)
        else:
            occupancy_data = []
            for date in date_range:
                occupied_count = db.session.query(func.count(RoomReservation.room_reservation_id)).filter(
                    cast(RoomReservation.room_reservation_booking_date_start, Date) <= date,
                    cast(RoomReservation.room_reservation_booking_date_end, Date) >= date
                ).scalar() or 0
                
                occupancy_data.append({
                    "date": date.strftime('%Y-%m-%d'),
                    "occupancy": occupied_count
                })

        # Revenue data
        if view_mode == 'monthly':
            revenue_data = db.session.query(
                func.DATE_FORMAT(Receipt.receipt_date, '%Y-%m').label('month'),
                func.sum(Receipt.receipt_total_amount).label('revenue')
            ).filter(
                cast(Receipt.receipt_date, Date) >= start_date,
                cast(Receipt.receipt_date, Date) <= end_date
            ).group_by('month').all()

            revenue_data = [
                {"date": item.month, "revenue": float(item.revenue or 0)}
                for item in revenue_data
            ]
        else:
            revenue_data = []
            for date in date_range:
                daily_revenue = db.session.query(func.sum(Receipt.receipt_total_amount)).filter(
                    cast(Receipt.receipt_date, Date) == date
                ).scalar() or 0
                
                revenue_data.append({
                    "date": date.strftime('%Y-%m-%d'),
                    "revenue": float(daily_revenue)
                })

        # Current occupied rooms
        current_occupied_rooms = db.session.query(func.count(RoomReservation.room_reservation_id)).filter(
            cast(RoomReservation.room_reservation_booking_date_start, Date) <= end_date,
            cast(RoomReservation.room_reservation_booking_date_end, Date) >= start_date
        ).scalar() or 0

        # Previous occupied rooms
        prev_occupied_rooms = db.session.query(func.count(RoomReservation.room_reservation_id)).filter(
            cast(RoomReservation.room_reservation_booking_date_start, Date) <= prev_end_date,
            cast(RoomReservation.room_reservation_booking_date_end, Date) >= prev_start_date
        ).scalar() or 0

        current_available_rooms = total_rooms - current_occupied_rooms
        prev_available_rooms = total_rooms - prev_occupied_rooms
        rooms_change = calculate_percentage_change(current_available_rooms, prev_available_rooms)

        # Room type performance
        room_type_performance = db.session.query(
            RoomType.room_type_name,
            func.count(RoomReservation.room_reservation_id).label('bookings'),
            func.sum(
                func.datediff(
                    RoomReservation.room_reservation_booking_date_end,
                    RoomReservation.room_reservation_booking_date_start
                )
            ).label('duration')
        ).join(
            Room, Room.room_type_id == RoomType.room_type_id
        ).join(
            RoomReservation, Room.room_id == RoomReservation.room_id
        ).filter(
            cast(RoomReservation.room_reservation_booking_date_start, Date) <= end_date,
            cast(RoomReservation.room_reservation_booking_date_end, Date) >= start_date
        ).group_by(RoomType.room_type_name).all()

        room_type_data = [
            {
                "roomType": row.room_type_name,
                "bookingFrequency": row.bookings,
                "avgStayDuration": float(row.duration / row.bookings if row.bookings > 0 else 0)
            }
            for row in room_type_performance
        ]

        # Visitor data with trending
        visitor_data = [
            {"name": "Room", "visitors": len(room_reservations)},
            {"name": "Venue", "visitors": len(venue_reservations)}
        ]

        total_visitors = len(room_reservations) + len(venue_reservations)
        prev_total_visitors = len(prev_room_reservations) + len(prev_venue_reservations)
        visitor_trending = calculate_percentage_change(total_visitors, prev_total_visitors)

        dashboard_data = {
            "totalBookings": current_bookings,
            "totalBookingsChange": round(calculate_percentage_change(current_bookings, prev_bookings), 1),
            "totalBookingsPeriod": "month" if view_mode == 'monthly' else "day",
            "totalRevenue": float(current_revenue),
            "totalRevenueChange": round(calculate_percentage_change(current_revenue, prev_revenue), 1),
            "totalRevenuePeriod": "month" if view_mode == 'monthly' else "day",
            "totalGuests": current_guests,
            "totalGuestsChange": round(calculate_percentage_change(current_guests, prev_guests), 1),
            "totalGuestsPeriod": "month" if view_mode == 'monthly' else "day",
            "availableRooms": current_available_rooms,
            "availableRoomsChange": round(rooms_change, 1),
            "availableRoomsPeriod": "month" if view_mode == 'monthly' else "day",
            "occupancyData": occupancy_data,
            "revenueData": revenue_data,
            "roomTypePerformance": room_type_data,
            "visitorData": visitor_data,
            "visitorTrending": round(visitor_trending, 1)
        }

        # Handle export format
        export_format = request.args.get('export')
        if export_format == 'json':
            return jsonify(dashboard_data)

        return jsonify(dashboard_data)

    except Exception as e:
        print(f"Error in get_dashboard_data: {str(e)}")
        return jsonify({"error": "An error occurred while fetching dashboard data"}), 500

