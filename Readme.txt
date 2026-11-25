QuickBite - Assessment.

In the excel sheet,it have  six tables such as areas,restaurants,delivery_partners,customers, orders, feedback.

Data count for each table:
i) area table - 21 records
ii) restaurant table - 44 records
iii) delivery_partners- 51 records
iv) customer table- 414 records
v) order table - 1567 records
vi) feedback table - 711 records

In each table it have attributes such as 
i) areas-area_id, area_name,city, zone
ii)restaurants- restaurant_id, restaurant_name, area_id, cusine, rating, onboard_date.
iii) delivery_partners- partner_id, partner_name, phone, vehicle_type, join_date
iv)customers- customer_id, customer_name, phone, email, area_id, signup_date
v) orders- order_id, customer_id, restarunt_id, partner_id, order_datetime, order_amount, payment_mode, order status,delivery_minutes, coupon_code
vi) feedback - feedback_id, order_id, rating, comment, feedback_date.

Key Identification in each table:
i) area- area_id (PK)
ii) restaurants- restaurant_id(PK), area_id(FK)
iii) delivery_partners- partner_id(PK)
iv) customers- customer_id(PK), area_id(FK)
v) orders- order_id(PK)
vi) feedbacks- feedback_id(PK)

Relationship between tables:
Areas → Restaurants: One area can have multiple restaurants.
Areas → Customers: One area can have multiple customers.
Customers → Orders: One customer can place multiple orders.
Restaurants → Orders: One restaurant can receive multiple orders.
Delivery Partners → Orders: One delivery partner can handle multiple orders.
Orders → Feedback: One order can have multiple feedback records, depending on the data.

Project Strcuture:

Task 1 - Cleaning the data
i) Duplicate record analysis for each tables
    1. area table - city,zone
    2. restaurant table - area_id, rating, cuisine
    3. delivery_partners table- vehicle_type
    4. customers table-area_id
    5. order table- customer_id, resturant_id, partner_id 
    6. feedback table- rating

ii) Contains null values for each table
    1. area table - no null values
    2. restaurant table -rating
    3. delivery_partners table- no null values
    4. customers table-area_id, email
    5. order table- customer_id, resturant_id, partner_id, deleviery_minutes, coupen_code
    6. feedback table- comments


Used Python and Pandas to clean all six datasets. 
Read the Excel sheets using pd.read_excel() and removed extra spaces using strip_text(). Standardized ID values using clean_id() and corrected inconsistent entries in city, cuisine, vehicle type, payment mode and order status using the standard() function with mapping dictionaries.
Removed duplicate records using drop_duplicates() through the remove_duplicates() function. 
Handled missing values using dropna() through the drop_missing() function. 
Cleaned phone numbers using the re module and clean_phone() function by removing unwanted characters. 
Converted date values using the custom to_date() function and converted order amount and delivery time into numeric values using Pandas conversion methods. 
Validated feedback ratings by converting them into numeric values and retaining only ratings between 1 and 5.
Saved all six cleaned datasets as separate sheets in Quick_Bite_Data_Cleaning.xlsx.

Task 2 – Data Combining
Used Python and Pandas merge() with LEFT JOIN to combine the cleaned Orders table with Customers, Restaurants, Areas, Delivery Partners and Feedback using common ID columns. 
Used sort_values() and drop_duplicates() to handle multiple feedback records. 
Generated a combined dataset and saved it as Quick_Bite_Combined_Data.xlsx.

Task 3 – Data Analysis
Used Python and Pandas functions such as groupby(), agg(), sum(), mean(), nunique() and sort_values() to analyse the combined dataset.
	1.Area Performance: Calculated total orders, revenue, average order value, delivery time, late rate and customer ratings.
	2.Top Cuisine by Day: Identified the most popular cuisine for each day using order counts and percentage share.
	3.Vehicle Performance: Compared vehicle types based on average delivery time, late orders and on-time delivery rate.
Saved the results in Quick_Bite_Analysis.xlsx.

Commands to Run:
i)Intial installation:                   pip install pandas openpyxl
                                         pip install numpy
ii)Run the file for cleaning:            python scripts/cleaning.py
iii)Run the file for analysing:          python scripts/analysis.py