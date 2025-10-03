Car Rental Service Application

📌 Introduction

The Car Rental Service Application is a small-scale management system built to streamline booking and rental operations for a car rental business. It integrates Microsoft Access, Excel, and VBA to automate and simplify key processes such as:

Booking and tracking car rentals

Monitoring vehicle availability

Calculating revenue

Managing customer details

This system minimizes manual work, reduces errors, and provides real-time insights into business performance.

🗂 Database Structure

The application is powered by an Access database with three core tables:

Customers Table

CustomerID (Primary Key)

CustomerName, Phone, Email, LicenseNumber, Address

Cars Table

CarID (Primary Key)

CarBrand, CarModel, CarType, Year, DailyRate, NumberPlate

Bookings Table

BookingID (Primary Key)

CustomerID, CarID, StartDate, EndDate, TotalCost

The Bookings Table links customers and cars using foreign keys.

🔍 Queries

The database includes queries to generate insights:

Average Rental Days Per Car – Calculates average rental duration by model.

Booking Count Per Car – Counts total rentals per model.

Detailed Booking Information – Retrieves full linked records.

Total Revenue by Car – Summarizes earnings per car model.

📊 Excel Front-End

The Excel interface makes the system easy to use with three sheets:

Records Sheet

Displays all booking data with conditional formatting:

🔴 Red: Booked

🟢 Green: Active

🔵 Blue: Completed

Load Data button refreshes from the database.

Form Sheet

Main interface for adding new bookings.

Fields for customer and booking details.

Features:

Auto-fetch customer info using VLOOKUP

Drop-down car brand/model selection

Automatic day count and cost calculation

Save Data button sends form data to the database.

Pivot Table Sheet

Summarizes monthly revenue per car.

Provides business insights for decision-making.

⚙️ VBA Middleware

The VBA layer connects Excel with Access to handle data and automation:

LoadBookingData – Pulls records into Excel, applies statuses (Booked, Active, Completed).

SaveNewBooking – Inserts new customer and booking records into Access.

DisplayCarInfo – Retrieves car details for pivot analysis.

✅ Features

Automated booking and record management

Real-time booking status tracking

Dynamic revenue reporting

Intuitive Excel-based UI for non-technical users

🚀 Future Enhancements

Maintenance Tracking – Log repairs and costs per vehicle.

Extra Cost Mechanism – Add hourly penalties for late returns.

Cloud Deployment – Scale for larger businesses.

Advanced Analytics – Integrate BI tools for deeper insights.
