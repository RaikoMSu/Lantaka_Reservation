from flask import Blueprint, jsonify, request, make_response
from datetime import datetime, timedelta, date
from sqlalchemy import cast, Date, func, and_, or_, extract
from model import db, RoomReservation, VenueReservation, Room, RoomType, Receipt, GuestDetails
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from calendar import monthrange
import logging
import matplotlib
matplotlib.use('Agg')  # Set the backend to Agg before importing pyplot
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.dates import DateFormatter

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

dashboard_bp = Blueprint('dashboard', __name__)

def get_month_range(date_obj):
    """Get the first and last day of a month."""
    first_day = date_obj.replace(day=1)
    _, last_day = monthrange(date_obj.year, date_obj.month)
    return first_day, date_obj.replace(day=last_day)

def format_currency(amount):
    """Format amount as PHP currency."""
    return f"₱{amount:,.2f}"

def calculate_percentage_change(current, previous):
    """Calculate percentage change with proper handling of edge cases."""
    try:
        if previous == 0:
            return 100 if current > 0 else 0
        return ((current - previous) / abs(previous)) * 100
    except (TypeError, ZeroDivisionError):
        return 0

def get_date_range_filter(start_date, end_date, date_column):
    """Create a date range filter for SQLAlchemy queries."""
    return and_(
        cast(date_column, Date) >= start_date,
        cast(date_column, Date) <= end_date
    )

def reset_date_range(start_date, end_date, view_mode):
    """Reset date range based on view mode."""
    if view_mode == 'monthly':
        # Reset to monthly range
        start_date = start_date.replace(day=1)
        _, last_day = monthrange(end_date.year, end_date.month)
        end_date = end_date.replace(day=last_day)
    elif view_mode == 'daily':
        # Reset to show current day and the last 7 days
        end_date = date.today()
        start_date = end_date - timedelta(days=6)
    return start_date, end_date

@dashboard_bp.route('/api/dashboardData', methods=['GET'])
def get_dashboard_data():
    try:
        # Parse and validate input parameters
        end_date_str = request.args.get('endDate')
        start_date_str = request.args.get('startDate')
        view_mode = request.args.get('viewMode', 'daily')
        export_format = request.args.get('export')

        try:
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date() if end_date_str else date.today()
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date() if start_date_str else (end_date - timedelta(days=6))
        except ValueError as e:
            return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400

        if start_date > end_date:
            return jsonify({"error": "Start date cannot be after end date"}), 400

        # Reset date range based on view mode
        start_date, end_date = reset_date_range(start_date, end_date, view_mode)

        logger.info(f"Processing request - Start Date: {start_date}, End Date: {end_date}, View Mode: {view_mode}")

        # Query reservations with optimized date filtering
        room_reservations = RoomReservation.query.filter(
            or_(
                and_(
                    get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start),
                    get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_end)
                )
            )
        ).all()

        venue_reservations = VenueReservation.query.filter(
            or_(
                and_(
                    get_date_range_filter(start_date, end_date, VenueReservation.venue_reservation_booking_date_start),
                    get_date_range_filter(start_date, end_date, VenueReservation.venue_reservation_booking_date_end)
                )
            )
        ).all()

        # Calculate metrics
        current_bookings = len(room_reservations) + len(venue_reservations)
        current_revenue = db.session.query(
            func.sum(Receipt.receipt_total_amount)
        ).filter(
            get_date_range_filter(start_date, end_date, Receipt.receipt_date)
        ).scalar() or 0

        # Calculate previous period
        period_length = (end_date - start_date).days + 1
        prev_end_date = start_date - timedelta(days=1)
        prev_start_date = prev_end_date - timedelta(days=period_length - 1)

        # Previous period metrics with aligned date ranges
        prev_room_reservations = RoomReservation.query.filter(
            get_date_range_filter(prev_start_date, prev_end_date, RoomReservation.room_reservation_booking_date_start)
        ).all()

        prev_venue_reservations = VenueReservation.query.filter(
            get_date_range_filter(prev_start_date, prev_end_date, VenueReservation.venue_reservation_booking_date_start)
        ).all()

        prev_bookings = len(prev_room_reservations) + len(prev_venue_reservations)
        prev_revenue = db.session.query(
            func.sum(Receipt.receipt_total_amount)
        ).filter(
            get_date_range_filter(prev_start_date, prev_end_date, Receipt.receipt_date)
        ).scalar() or 0

        # Generate time series data with aligned dates
        dates = []
        if view_mode == 'monthly':
            current_date = start_date
            while current_date <= end_date:
                dates.append(current_date)
                if current_date.month == 12:
                    current_date = current_date.replace(year=current_date.year + 1, month=1, day=1)
                else:
                    current_date = current_date.replace(month=current_date.month + 1, day=1)
        else:
            dates = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]

        # Prepare time series data with consistent formatting
        occupancy_data = []
        revenue_data = []

        for current_date in dates:
            if view_mode == 'monthly':
                month_start, month_end = get_month_range(current_date)
                
                # Monthly metrics with consistent date ranges
                occupied_count = db.session.query(func.count(RoomReservation.room_id)).filter(
                    and_(
                        cast(RoomReservation.room_reservation_booking_date_start, Date) <= month_end,
                        cast(RoomReservation.room_reservation_booking_date_end, Date) >= month_start
                    )
                ).scalar() or 0

                monthly_revenue = db.session.query(
                    func.sum(Receipt.receipt_total_amount)
                ).filter(
                    and_(
                        cast(Receipt.receipt_date, Date) >= month_start,
                        cast(Receipt.receipt_date, Date) <= month_end
                    )
                ).scalar() or 0

                formatted_date = current_date.strftime('%Y-%m')
                
                occupancy_data.append({
                    "date": formatted_date,
                    "occupancy": occupied_count
                })
                
                revenue_data.append({
                    "date": formatted_date,
                    "revenue": float(monthly_revenue)
                })
            else:
                # Daily metrics with consistent date format
                occupied_count = db.session.query(func.count(RoomReservation.room_id)).filter(
                    and_(
                        cast(RoomReservation.room_reservation_booking_date_start, Date) <= current_date,
                        cast(RoomReservation.room_reservation_booking_date_end, Date) >= current_date
                    )
                ).scalar() or 0

                daily_revenue = db.session.query(
                    func.sum(Receipt.receipt_total_amount)
                ).filter(
                    cast(Receipt.receipt_date, Date) == current_date
                ).scalar() or 0

                formatted_date = current_date.strftime('%Y-%m-%d')
                
                occupancy_data.append({
                    "date": formatted_date,
                    "occupancy": occupied_count
                })
                
                revenue_data.append({
                    "date": formatted_date,
                    "revenue": float(daily_revenue)
                })

        # Room type performance
        room_type_performance = db.session.query(
            RoomType.room_type_name,
            func.count(RoomReservation.room_reservation_id).label('bookings'),
            func.avg(
                func.datediff(
                    RoomReservation.room_reservation_booking_date_end,
                    RoomReservation.room_reservation_booking_date_start
                )
            ).label('avg_duration')
        ).join(
            Room, Room.room_type_id == RoomType.room_type_id
        ).join(
            RoomReservation, Room.room_id == RoomReservation.room_id
        ).filter(
            get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start)
        ).group_by(RoomType.room_type_name).all()

        room_type_data = [
            {
                "roomType": row.room_type_name,
                "bookingFrequency": row.bookings,
                "avgStayDuration": float(row.avg_duration or 0)
            }
            for row in room_type_performance
        ]

        # Visitor data with specific colors
        visitor_data = [
            {
                "name": "Room Guests",
                "visitors": len(room_reservations),
            },
            {
                "name": "Venue Visitors",
                "visitors": len(venue_reservations),
            }
        ]

        total_visitors = len(room_reservations) + len(venue_reservations)
        prev_total_visitors = len(prev_room_reservations) + len(prev_venue_reservations)
        visitor_trending = calculate_percentage_change(total_visitors, prev_total_visitors)

        # Room availability calculations
        total_rooms = db.session.query(func.count(Room.room_id)).scalar() or 0
        occupied_rooms = db.session.query(func.count(RoomReservation.room_id)).filter(
            get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start)
        ).scalar() or 0
        
        available_rooms = total_rooms - occupied_rooms

        # Guest metrics
        total_guests = db.session.query(
            func.count(func.distinct(GuestDetails.guest_id))
        ).filter(
            or_(
                GuestDetails.guest_id.in_([r.guest_id for r in room_reservations]),
                GuestDetails.guest_id.in_([v.guest_id for v in venue_reservations])
            )
        ).scalar() or 0

        # Previous period calculations
        prev_occupied_rooms = db.session.query(
            func.count(RoomReservation.room_id)
        ).filter(
            get_date_range_filter(prev_start_date, prev_end_date, RoomReservation.room_reservation_booking_date_start)
        ).scalar() or 0
        
        prev_available_rooms = total_rooms - prev_occupied_rooms
        
        prev_total_guests = db.session.query(
            func.count(func.distinct(GuestDetails.guest_id))
        ).filter(
            or_(
                GuestDetails.guest_id.in_([r.guest_id for r in prev_room_reservations]),
                GuestDetails.guest_id.in_([v.guest_id for v in prev_venue_reservations])
            )
        ).scalar() or 0

        # Prepare dashboard response
        dashboard_data = {
            "totalBookings": current_bookings,
            "totalBookingsChange": round(calculate_percentage_change(current_bookings, prev_bookings), 1),
            "totalBookingsPeriod": "month" if view_mode == 'monthly' else "week",
            "totalRevenue": float(current_revenue),
            "totalRevenueChange": round(calculate_percentage_change(current_revenue, prev_revenue), 1),
            "totalRevenuePeriod": "month" if view_mode == 'monthly' else "week",
            "occupancyData": occupancy_data,
            "revenueData": revenue_data,
            "roomTypePerformance": room_type_data,
            "visitorData": visitor_data,
            "visitorTrending": round(visitor_trending, 1),
            "availableRooms": available_rooms,
            "availableRoomsChange": round(calculate_percentage_change(available_rooms, prev_available_rooms), 1),
            "totalGuests": total_guests,
            "totalGuestsChange": round(calculate_percentage_change(total_guests, prev_total_guests), 1)
        }

        # Handle export formats
        if export_format:
            if export_format == 'excel':
                return export_excel(dashboard_data, occupancy_data, revenue_data, room_type_data, visitor_data)
            elif export_format == 'pdf':
                return export_pdf(dashboard_data)
            else:
                return jsonify({"error": "Unsupported export format"}), 400

        return jsonify(dashboard_data)

    except Exception as e:
        logger.error(f"Error in get_dashboard_data: {str(e)}", exc_info=True)
        return jsonify({"error": "An internal server error occurred"}), 500

def export_excel(dashboard_data, occupancy_data, revenue_data, room_type_data, visitor_data):
    """Generate Excel report."""
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Add formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4B5563',
                'font_color': 'white',
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'border': 1
            })

            # Summary sheet
            summary_df = pd.DataFrame({
                'Metric': [
                    'Total Bookings',
                    'Total Revenue',
                    'Available Rooms',
                    'Total Guests',
                    'Visitor Trend'
                ],
                'Current Value': [
                    dashboard_data['totalBookings'],
                    format_currency(dashboard_data['totalRevenue']),
                    dashboard_data['availableRooms'],
                    dashboard_data['totalGuests'],
                    f"{dashboard_data['visitorTrending']}%"
                ],
                'Change': [
                    f"{dashboard_data['totalBookingsChange']}%",
                    f"{dashboard_data['totalRevenueChange']}%",
                    f"{dashboard_data['availableRoomsChange']}%",
                    f"{dashboard_data['totalGuestsChange']}%",
                    "N/A"
                ]
            })
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Format Summary sheet
            summary_sheet = writer.sheets['Summary']
            for col_num, value in enumerate(summary_df.columns.values):
                summary_sheet.write(0, col_num, value, header_format)

            # Occupancy sheet
            occupancy_df = pd.DataFrame(occupancy_data)
            occupancy_df.to_excel(writer, sheet_name='Occupancy', index=False)
            
            # Revenue sheet
            revenue_df = pd.DataFrame(revenue_data)
            revenue_df['revenue'] = revenue_df['revenue'].apply(lambda x: format_currency(x))
            revenue_df.to_excel(writer, sheet_name='Revenue', index=False)

            # Room Performance sheet
            room_perf_df = pd.DataFrame(room_type_data)
            room_perf_df.columns = ['Room Type', 'Booking Frequency', 'Average Stay Duration (Days)']
            room_perf_df.to_excel(writer, sheet_name='Room Performance', index=False)

            # Visitors sheet with colors
            visitors_df = pd.DataFrame(visitor_data)
            visitors_df.to_excel(writer, sheet_name='Visitors', index=False)

            # Auto-adjust columns width
            for sheet in writer.sheets.values():
                sheet.autofit()

        output.seek(0)
        response = make_response(output.getvalue())
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response.headers['Content-Disposition'] = f'attachment; filename=dashboard-report-{datetime.now().strftime("%Y-%m-%d")}.xlsx'
        return response

    except Exception as e:
        logger.error(f"Error in export_excel: {str(e)}", exc_info=True)
        return jsonify({"error": "Failed to generate Excel report"}), 500

def export_pdf(dashboard_data):
    """Generate PDF report with charts."""
    try:
        import matplotlib.pyplot as plt
        import seaborn as sns
        from matplotlib.dates import DateFormatter
        import pandas as pd
        from datetime import datetime

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []

        # Styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=24,
            spaceAfter=30
        )
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Normal'],
            fontSize=12,
            textColor=colors.gray
        )
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12
        )

        # Title and timestamp
        elements.append(Paragraph("Dashboard Report", title_style))
        elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_style))
        elements.append(Spacer(1, 20))

        # Summary table
        elements.append(Paragraph("Summary Metrics", header_style))
        
        summary_data = [
            ['Metric', 'Current Value', 'Change vs Previous Period'],
            ['Total Bookings', str(dashboard_data['totalBookings']), f"{dashboard_data['totalBookingsChange']}%"],
            ['Total Revenue', format_currency(dashboard_data['totalRevenue']), f"{dashboard_data['totalRevenueChange']}%"],
            ['Available Rooms', str(dashboard_data['availableRooms']), f"{dashboard_data['availableRoomsChange']}%"],
            ['Total Guests', str(dashboard_data['totalGuests']), f"{dashboard_data['totalGuestsChange']}%"],
        ]

        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B5563')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWHEIGHT', (0, 0), (-1, -1), 25)
        ])

        summary_table = Table(summary_data, colWidths=[2.5*inch, 2.5*inch, 2.5*inch])
        summary_table.setStyle(table_style)
        elements.append(summary_table)
        elements.append(Spacer(1, 20))

        # Create charts
        # 1. Occupancy Chart
        elements.append(Paragraph("Occupancy Trends", header_style))
        plt.figure(figsize=(10, 4))
        occupancy_df = pd.DataFrame(dashboard_data['occupancyData'])
        occupancy_df['date'] = pd.to_datetime(occupancy_df['date'])
        plt.plot(occupancy_df['date'], occupancy_df['occupancy'], marker='o')
        plt.title('Room Occupancy Over Time')
        plt.xlabel('Date')
        plt.ylabel('Occupancy')
        plt.grid(True)
        plt.xticks(rotation=45)
        
        # Save chart to memory
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        elements.append(Image(img_buffer, width=7*inch, height=3*inch))
        elements.append(Spacer(1, 20))
        plt.close()

        # 2. Revenue Chart
        elements.append(Paragraph("Revenue Analysis", header_style))
        plt.figure(figsize=(10, 4))
        revenue_df = pd.DataFrame(dashboard_data['revenueData'])
        revenue_df['date'] = pd.to_datetime(revenue_df['date'])
        plt.plot(revenue_df['date'], revenue_df['revenue'], marker='o', color='green')
        plt.title('Revenue Trends')
        plt.xlabel('Date')
        plt.ylabel('Revenue (PHP)')
        plt.grid(True)
        plt.xticks(rotation=45)
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        elements.append(Image(img_buffer, width=7*inch, height=3*inch))
        elements.append(Spacer(1, 20))
        plt.close()

        # 3. Visitor Distribution Pie Chart
        elements.append(Paragraph("Visitor Distribution", header_style))
        plt.figure(figsize=(8, 8))
        visitor_df = pd.DataFrame(dashboard_data['visitorData'])
        plt.pie(visitor_df['visitors'], labels=visitor_df['name'], autopct='%1.1f%%')
        plt.title('Distribution of Visitors')
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        elements.append(Image(img_buffer, width=5*inch, height=5*inch))
        elements.append(Spacer(1, 20))
        plt.close()

        # 4. Room Type Performance
        elements.append(Paragraph("Room Type Performance", header_style))
        room_perf_df = pd.DataFrame(dashboard_data['roomTypePerformance'])
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4))
        
        # Booking Frequency
        sns.barplot(data=room_perf_df, x='roomType', y='bookingFrequency', ax=ax1)
        ax1.set_title('Booking Frequency by Room Type')
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45)
        
        # Average Stay Duration
        sns.barplot(data=room_perf_df, x='roomType', y='avgStayDuration', ax=ax2)
        ax2.set_title('Average Stay Duration by Room Type')
        ax2.set_xticklabels(ax2.get_xticklabels(), rotation=45)
        
        plt.tight_layout()
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        elements.append(Image(img_buffer, width=7*inch, height=3*inch))
        plt.close()

        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        response = make_response(buffer.getvalue())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=dashboard-report-{datetime.now().strftime("%Y-%m-%d")}.pdf'
        return response

    except Exception as e:
        logger.error(f"Error in export_pdf: {str(e)}", exc_info=True)
        return jsonify({"error": "Failed to generate PDF report"}), 500