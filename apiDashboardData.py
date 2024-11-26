from flask import Blueprint, jsonify, request, make_response
from datetime import datetime, timedelta, date
from sqlalchemy import cast, Date, func, and_, or_, extract
from model import db, RoomReservation, VenueReservation, Room, RoomType, Receipt, GuestDetails, Venue
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

def get_available_spaces(start_date, end_date):
    """Calculate available spaces considering ready status only."""
    # Get total rooms and venues that are ready
    total_rooms = db.session.query(func.count(Room.room_id))\
        .filter(Room.room_status == 'ready')\
        .scalar() or 0
    
    total_venues = db.session.query(func.count(Venue.venue_id))\
        .filter(Venue.venue_status == 'ready')\
        .scalar() or 0
    
    # Get occupied rooms and venues in date range
    occupied_rooms = db.session.query(func.count(RoomReservation.room_id))\
        .join(Room)\
        .filter(
            and_(
                Room.room_status == 'ready',
                get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start)
            )
        ).scalar() or 0
    
    occupied_venues = db.session.query(func.count(VenueReservation.venue_id))\
        .join(Venue)\
        .filter(
            and_(
                Venue.venue_status == 'ready',
                get_date_range_filter(start_date, end_date, VenueReservation.venue_reservation_booking_date_start)
            )
        ).scalar() or 0
    
    return (total_rooms - occupied_rooms) + (total_venues - occupied_venues)

def reset_date_range(start_date, end_date, view_mode):
    """Reset date range based on view mode."""
    if view_mode == 'monthly':
        # Reset to monthly range
        start_date = start_date.replace(day=1)
        _, last_day = monthrange(end_date.year, end_date.month)
        end_date = end_date.replace(day=last_day)
    else:
        # For daily view, show last 7 days
        end_date = date.today()
        start_date = end_date - timedelta(days=6)
    return start_date, end_date

@dashboard_bp.route('/api/dashboardData', methods=['GET'])
def get_dashboard_data():
    try:
        # Parse and validate input parameters
        end_date_str = request.args.get('endDate')
        start_date_str = request.args.get('startDate')
        view_mode = request.args.get('viewMode', 'monthly')
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
        room_reservations = RoomReservation.query.join(Room).filter(
            and_(
                Room.room_status == 'ready',
                get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start)
            )
        ).all()

        venue_reservations = VenueReservation.query.join(Venue).filter(
            and_(
                Venue.venue_status == 'ready',
                get_date_range_filter(start_date, end_date, VenueReservation.venue_reservation_booking_date_start)
            )
        ).all()

        # Calculate metrics
        current_bookings = len(room_reservations) + len(venue_reservations)
        current_revenue = db.session.query(
            func.sum(Receipt.receipt_total_amount)
        ).filter(
            get_date_range_filter(start_date, end_date, Receipt.receipt_date)
        ).scalar() or 0

        # Calculate available spaces
        current_available_spaces = get_available_spaces(start_date, end_date)

        # Calculate previous period
        period_length = (end_date - start_date).days + 1
        prev_end_date = start_date - timedelta(days=1)
        prev_start_date = prev_end_date - timedelta(days=period_length - 1)

        # Previous period metrics
        prev_room_reservations = RoomReservation.query.join(Room).filter(
            and_(
                Room.room_status == 'ready',
                get_date_range_filter(prev_start_date, prev_end_date, RoomReservation.room_reservation_booking_date_start)
            )
        ).all()

        prev_venue_reservations = VenueReservation.query.join(Venue).filter(
            and_(
                Venue.venue_status == 'ready',
                get_date_range_filter(prev_start_date, prev_end_date, VenueReservation.venue_reservation_booking_date_start)
            )
        ).all()

        prev_bookings = len(prev_room_reservations) + len(prev_venue_reservations)
        prev_revenue = db.session.query(
            func.sum(Receipt.receipt_total_amount)
        ).filter(
            get_date_range_filter(prev_start_date, prev_end_date, Receipt.receipt_date)
        ).scalar() or 0

        prev_available_spaces = get_available_spaces(prev_start_date, prev_end_date)

        # Generate time series data
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

        # Prepare time series data
        occupancy_data = []
        revenue_data = []

        for current_date in dates:
            if view_mode == 'monthly':
                month_start, month_end = get_month_range(current_date)
                
                occupied_count = db.session.query(func.count(RoomReservation.room_id))\
                    .join(Room)\
                    .filter(
                        and_(
                            Room.room_status == 'ready',
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
                occupied_count = db.session.query(func.count(RoomReservation.room_id))\
                    .join(Room)\
                    .filter(
                        and_(
                            Room.room_status == 'ready',
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
            and_(
                Room.room_status == 'ready',
                get_date_range_filter(start_date, end_date, RoomReservation.room_reservation_booking_date_start)
            )
        ).group_by(RoomType.room_type_name).all()

        room_type_data = [
            {
                "roomType": row.room_type_name,
                "bookingFrequency": row.bookings,
                "avgStayDuration": float(row.avg_duration or 0)
            }
            for row in room_type_performance
        ]

        # Visitor data
        visitor_data = [
            {
                "name": "Room Guests",
                "visitors": len(room_reservations)
            },
            {
                "name": "Venue Visitors",
                "visitors": len(venue_reservations)
            }
        ]

        total_visitors = len(room_reservations) + len(venue_reservations)
        prev_total_visitors = len(prev_room_reservations) + len(prev_venue_reservations)
        visitor_trending = calculate_percentage_change(total_visitors, prev_total_visitors)

        # Guest metrics
        total_guests = db.session.query(
            func.count(func.distinct(GuestDetails.guest_id))
        ).filter(
            or_(
                GuestDetails.guest_id.in_([r.guest_id for r in room_reservations]),
                GuestDetails.guest_id.in_([v.guest_id for v in venue_reservations])
            )
        ).scalar() or 0

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
            "availableSpaces": current_available_spaces,
            "availableSpacesChange": round(calculate_percentage_change(current_available_spaces, prev_available_spaces), 1),
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
                    'Available Spaces',
                    'Total Guests',
                    'Visitor Trend'
                ],
                'Current Value': [
                    dashboard_data['totalBookings'],
                    format_currency(dashboard_data['totalRevenue']),
                    dashboard_data['availableSpaces'],
                    dashboard_data['totalGuests'],
                    f"{dashboard_data['visitorTrending']}%"
                ],
                'Change': [
                    f"{dashboard_data['totalBookingsChange']}%",
                    f"{dashboard_data['totalRevenueChange']}%",
                    f"{dashboard_data['availableSpacesChange']}%",
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

            # Visitors sheet
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
    """Generate enhanced PDF report with tables and charts."""
    try:
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30
        )
        elements = []

        # Styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=24,
            spaceAfter=30,
            textColor=colors.HexColor('#1a365d')
        )
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=20,
            textColor=colors.HexColor('#2c5282')
        )
        subheader_style = ParagraphStyle(
            'CustomSubHeader',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=10,
            textColor=colors.HexColor('#4a5568')
        )
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['Normal'],
            fontSize=10,
            leading=14,
            spaceAfter=12
        )

        # Title and Date
        elements.append(Paragraph("Hotel Performance Dashboard Report", title_style))
        elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subheader_style))
        elements.append(Spacer(1, 20))

        # Summary Table
        elements.append(Paragraph("Performance Summary", header_style))
        
        summary_data = [
            ['Metric', 'Current Value', 'Change'],
            ['Total Bookings', str(dashboard_data['totalBookings']), f"{dashboard_data['totalBookingsChange']}%"],
            ['Total Revenue', format_currency(dashboard_data['totalRevenue']), f"{dashboard_data['totalRevenueChange']}%"],
            ['Available Spaces', str(dashboard_data['availableSpaces']), f"{dashboard_data['availableSpacesChange']}%"],
            ['Total Guests', str(dashboard_data['totalGuests']), f"{dashboard_data['totalGuestsChange']}%"]
        ]

        summary_table = Table(summary_data, colWidths=[200, 150, 100])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a365d')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 30))

        # Occupancy Chart
        elements.append(Paragraph("Occupancy Trends", header_style))
        elements.append(Paragraph("""
            This graph illustrates the monthly occupancy trends of our facilities. The data demonstrates 
            the utilization of our spaces over time, providing insights into seasonal patterns and overall 
            demand fluctuations.
        """, body_style))
        
        plt.figure(figsize=(10, 6))
        occupancy_data = dashboard_data['occupancyData']
        
        dates = [d['date'] for d in occupancy_data]
        occupancy = [d['occupancy'] for d in occupancy_data]
        
        plt.bar(dates, occupancy, color='#3B82F6')
        plt.title('Monthly Occupancy')
        plt.xlabel('Month')
        plt.ylabel('Occupancy Count')
        plt.xticks(rotation=45)
        plt.grid(True, axis='y', linestyle='--', alpha=0.7)
        
        occupancy_buffer = BytesIO()
        plt.savefig(occupancy_buffer, format='png', bbox_inches='tight', dpi=300)
        occupancy_buffer.seek(0)
        elements.append(Image(occupancy_buffer, width=7*inch, height=4*inch))
        plt.close()

        elements.append(Spacer(1, 20))

        # Revenue Chart
        elements.append(Paragraph("Revenue Analysis", header_style))
        elements.append(Paragraph("""
            This graph presents our revenue performance over time. The curve demonstrates the financial 
            trajectory of our operations, allowing us to identify trends, peak periods, and areas for 
            potential growth or improvement in our revenue management strategies.
        """, body_style))
        
        plt.figure(figsize=(10, 6))
        revenue_data = dashboard_data['revenueData']
        
        dates = [d['date'] for d in revenue_data]
        revenue = [d['revenue'] for d in revenue_data]
        
        plt.plot(dates, revenue, color='#3B82F6', linewidth=2)
        plt.fill_between(dates, revenue, alpha=0.2, color='#3B82F6')
        plt.title('Revenue Trends')
        plt.xlabel('Month')
        plt.ylabel('Revenue (₱)')
        plt.xticks(rotation=45)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Format y-axis as currency
        plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₱{x:,.0f}'))
        
        revenue_buffer = BytesIO()
        plt.savefig(revenue_buffer, format='png', bbox_inches='tight', dpi=300)
        revenue_buffer.seek(0)
        elements.append(Image(revenue_buffer, width=7*inch, height=4*inch))
        plt.close()

        elements.append(Spacer(1, 20))

        # Visitor Distribution Chart
        elements.append(Paragraph("Visitor Distribution Analysis", header_style))
        elements.append(Paragraph(f"""
            This chart illustrates the distribution of visitors between room guests and venue visitors. 
            It provides a clear visualization of our customer segmentation, helping us understand the 
            balance between different types of visitors and inform our marketing and operational strategies.
        """, body_style))
        
        plt.figure(figsize=(8, 8))
        visitor_data = dashboard_data['visitorData']
        visitors = [d['visitors'] for d in visitor_data]
        labels = [d['name'] for d in visitor_data]
        colors_pie = ['#60A5FA', '#3B82F6']
        
        plt.pie(visitors,
                labels=labels,
                colors=colors_pie,
                autopct='%1.1f%%',
                startangle=90)
        plt.title('Room vs Venue Visitors')
        
        # Add total visitors in center
        total_visitors = sum(visitors)
        plt.text(0, 0, f'{total_visitors}\nTotal Visitors',
                ha='center', va='center',
                fontsize=12, fontweight='bold')
        
        visitor_buffer = BytesIO()
        plt.savefig(visitor_buffer, format='png', bbox_inches='tight', dpi=300)
        visitor_buffer.seek(0)
        elements.append(Image(visitor_buffer, width=4*inch, height=4*inch))
        plt.close()

        elements.append(Spacer(1, 20))

        # Room Type Performance Chart
        elements.append(Paragraph("Room Type Performance", header_style))
        elements.append(Paragraph("""
            This combined chart displays both booking frequency and average stay duration for each room type. 
            The bars represent the number of bookings, while the line shows the average length of stay. 
            This visualization helps identify our most popular room types and understand guest preferences 
            in terms of stay duration, enabling better inventory management and pricing strategies.
        """, body_style))
        
        plt.figure(figsize=(10, 6))
        room_data = dashboard_data['roomTypePerformance']
        
        fig, ax1 = plt.subplots(figsize=(10, 6))
        
        x = range(len(room_data))
        bars = ax1.bar(x, [d['bookingFrequency'] for d in room_data],
                      color='#60A5FA', alpha=0.7)
        ax1.set_ylabel('Booking Frequency', color='#60A5FA')
        ax1.tick_params(axis='y', labelcolor='#60A5FA')
        
        ax2 = ax1.twinx()
        line = ax2.plot(x, [d['avgStayDuration'] for d in room_data],
                       color='#3B82F6', marker='o', linewidth=2,
                       label='Avg Stay Duration')
        ax2.set_ylabel('Average Stay Duration (days)', color='#3B82F6')
        ax2.tick_params(axis='y', labelcolor='#3B82F6')
        
        plt.title('Booking Frequency and Average Stay Duration by Room Type')
        plt.xticks(x, [d['roomType'] for d in room_data], rotation=45)
        
        # Add legend
        lines_1, labels_1 = ax1.get_legend_handles_labels()
        lines_2, labels_2 = ax2.get_legend_handles_labels()
        ax2.legend(lines_1 + lines_2, ['Booking Frequency', 'Avg Stay Duration'],
                  loc='upper right')
        
        plt.tight_layout()
        
        performance_buffer = BytesIO()
        plt.savefig(performance_buffer, format='png', bbox_inches='tight', dpi=300)
        performance_buffer.seek(0)
        elements.append(Image(performance_buffer, width=7*inch, height=4*inch))
        plt.close()

        # Conclusion
        elements.append(Paragraph("Conclusion", header_style))
        elements.append(Paragraph(f"""
            This report provides a comprehensive overview of our hotel's performance. Key highlights include:
            
            1. Total Bookings: {dashboard_data['totalBookings']} ({dashboard_data['totalBookingsChange']}% change)
            2. Total Revenue: {format_currency(dashboard_data['totalRevenue'])} ({dashboard_data['totalRevenueChange']}% change)
            3. Available Spaces: {dashboard_data['availableSpaces']} ({dashboard_data['availableSpacesChange']}% change)
            4. Total Guests: {dashboard_data['totalGuests']} ({dashboard_data['totalGuestsChange']}% change)
            5. Visitor Distribution: {round(dashboard_data['visitorData'][0]['visitors']/sum(v['visitors'] for v in dashboard_data['visitorData'])*100)}% room guests, {round(dashboard_data['visitorData'][1]['visitors']/sum(v['visitors'] for v in dashboard_data['visitorData'])*100)}% venue visitors
            6. Visitor Trend: {dashboard_data['visitorTrending']}% increase this period
            
            These metrics and visualizations provide valuable insights into our operations, helping guide 
            strategic decisions in areas such as pricing, marketing, and capacity management.
        """, body_style))

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