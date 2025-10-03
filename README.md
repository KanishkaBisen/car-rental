Car Rental Service Application

📌 Introduction
```
The Car Rental Service Application is a small-scale management system built to streamline booking and rental operations for a car rental business. It integrates Microsoft Access, Excel, and VBA to automate and simplify key processes such as:

Booking and tracking car rentals
Monitoring vehicle availability
Calculating revenue
Managing customer details
```
This system minimizes manual work, reduces errors, and provides real-time insights into business performance.

🗂 Database Structure

| **Table**     | **Primary Key** | **Fields**                                                 |
| ------------- | --------------- | ---------------------------------------------------------- |
| **Customers** | `CustomerID`    | CustomerName, Phone, Email, LicenseNumber, Address         |
| **Cars**      | `CarID`         | CarBrand, CarModel, CarType, Year, DailyRate, NumberPlate  |
| **Bookings**  | `BookingID`     | CustomerID (FK), CarID (FK), StartDate, EndDate, TotalCost |

🔗 The Bookings Table links customers and cars via foreign keys (CustomerID, CarID).

🔍 Queries

The database includes queries to generate insights:

Average Rental Days Per Car – Calculates average rental duration by model.

Booking Count Per Car – Counts total rentals per model.

Detailed Booking Information – Retrieves full linked records.

Total Revenue by Car – Summarizes earnings per car model.

📊 Excel Front-End

Records Sheet

Displays all bookings with statuses (Booked 🔴, Active 🟢, Completed 🔵).

Load Data button refreshes data from the database.

<img width="363" height="194" alt="image" src="https://github.com/user-attachments/assets/fb0d11f1-7b68-40cb-9bf6-1c2d208ec008" />

Form Sheet

Interface to add/edit bookings.

Auto-fetches customer/car details with VLOOKUP.

Calculates total cost and duration automatically.

Save Data button submits booking data to the database.

Pivot Table Sheet

<img width="274" height="285" alt="image" src="https://github.com/user-attachments/assets/da9beccc-791b-46f8-ab6a-1374e432e117" />

Summarizes total revenue by car and time period.

Supports business insights and decision-making.

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
