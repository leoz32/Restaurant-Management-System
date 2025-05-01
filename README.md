# Restaurant Management System

This is a Restaurant Management Information System (MIS) web application developed based on the Flask framework.

## Main Features

*   **User Authentication:** Staff login and logout.
*   **Dashboard:** Displays an overview of key business indicators, such as total customers, total orders, total inventory, today's orders, recent orders, and low stock warnings.
*   **Customer Management:** View, add, edit, and delete customer information.
*   **Menu Management:** View, add, edit, and delete menu item information.
*   **Inventory Management:** View, add, edit, and delete inventory item information, linking suppliers and items.
*   **Order Management:** View, create, and manage customer orders.
*   **Staff Management:** View, add, edit, and delete staff information.
*   **Report Generation:**
    *   Sales Report
    *   Customer Report
    *   Menu Item Report
    *   Inventory Report
    *   Staff Report
    *   Financial Overview
    *   Seasonal Sales Analysis
    *   Food Waste Analysis
*   **Data Prediction:** Simple prediction of sales and customer numbers based on historical data.
*   **Data Management:** 

## Technology Stack

*   **Backend:** Python, Flask
*   **Database:** Microsoft Access (`.accdb`)
*   **Database Connection:** pyodbc
*   **Frontend:** HTML, CSS, JavaScript (Jinja2 template engine)
*   **Data Processing:** Pandas, NumPy (for prediction functionality)
*   **Caching:** Flask-Caching (SimpleCache)

## Environment Setup

1.  **Prerequisites:**
    *   **Python:** Ensure Python 3.x is installed.
    *   **Microsoft Access Database Engine:** You need to install the Microsoft Access Database Engine that matches your Python interpreter's architecture (32-bit or 64-bit). You can download it from the [Microsoft official website](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (please select the correct version). **Note:** If you have 32-bit Office/Access installed, you usually need to install 32-bit Python and the 32-bit Access Database Engine to connect via pyodbc.

2.  **Clone the Repository:**
    ```bash
    git clone <repository-url>
    cd flask-app-trae
    ```

3.  **Install Dependencies:**
    In the project root directory (containing `requirements.txt`), run:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Database:**
    Ensure the `Margaretds.accdb` database file is located in the project's root directory (at the same level as the `TESY.py` file).

## Running the Application

1.  Open a terminal or command prompt.
2.  Navigate to the project root directory.
3.  Run the following command to start the Flask development server:
    ```bash
    python TESY.py
    ```
    Or specify a port (e.g., 8080):
    ```bash
    python TESY.py --port 8080
    ```
4.  Once the server starts, access `http://127.0.0.1:5000` (or your specified port) in your browser.
5.  Default login credentials (based on `TESY.py` code): Username is `wei.chen@restaurant.com` from the Staff table, password is `admin`.

## Project Structure

```
flask-app-trae/
├── Margaretds.accdb       # Access database file
├── TESY.py                # Main Flask application file
├── requirements.txt       # Python dependency list
├── static/                # Static files (CSS, JavaScript, Images)
│   ├── css/
│   ├── js/
│   └── images/
├── templates/             # HTML template files (Jinja2)
│   ├── base.html          # Base template
│   ├── login.html         # Login page
│   ├── dashboard.html     # Dashboard page
│   ├── customers/         # Customer management related templates
│   ├── menu/              # Menu management related templates
│   ├── inventory/         # Inventory management related templates
│   ├── orders/            # Order management related templates
│   ├── staff/             # Staff management related templates
│   ├── reports/           # Report related templates
│   └── ...                # Other template files
└── ...                    # Other configuration files or directories
```
