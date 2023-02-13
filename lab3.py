from sys import argv
import sys
import os
from datetime import date
import pandas as pd
import re



def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    # Check whether provide parameter is valid path of file
    num_parameters = len(argv)-1
    if num_parameters >=1:
        csv_path = argv[1]

        if os.path.isfile(csv_path):
            return os.path.abspath(csv_path)
        else:
            print("Error found: The Csv do not exist check again.")
            sys.exit(1)
    else:
        print("Error found: Csv file path not found.")
        sys.exit(1)
    return

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_directory= os.path.dirname(sales_csv)
    # Determine the name and path of the directory to hold the order data files
    todays_date= date.today().isoformat()
    orders_dir_name= f'Orders_{todays_date}'
    orders_dir_path = os.path.join(sales_directory, orders_dir_name)
    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path


# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df= pd.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY']*sales_df['ITEM PRICE'])
    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS','CITY','STATE','POSTAL CODE','COUNTRY'], inplace=True)
    # Group the rows in the DataFrame by order ID
    for orders_id, order_df in sales_df.groupby('ORDER ID'):
    # For each order ID:
        # Remove the "ORDER ID" column
        order_df.drop(columns=['ORDER ID'], inplace=True)
        # Sort the items by item number
        order_df.sort_values(by=['ITEM NUMBER'], inplace=True)
        # Append a "GRAND TOTAL" row
        grand_total= order_df['TOTAL PRICE'].sum()
        grand_total_df= pd.DataFrame({'ITEM PRICE':['GRAND TOTAL:'],'TOTAL PRICE':[grand_total]})
        order_df= pd.concat([order_df,grand_total_df])
        # Determine the file name and full path of the Excel sheet
        customer_name= order_df['CUSTOMER NAME'].values[0]
        customer_name= re.sub(r'\W','',customer_name)
        orders_file_name=f'Order{orders_id}_{customer_name}.xlsx'
        orders_file_path=os.path.join(orders_dir,orders_file_name)
        # Export the data to an Excel sheet
        sheet_name=f'Order {orders_id}'
        order_df.to_excel(orders_file_path, index=False,sheet_name=sheet_name, )
        # TODO: Format the Excel sheet
        workbook  = orders_file_path.book
        worksheet = orders_file_path.sheets['Sheet1']
        sheet_name.set_column(2, 2, None, format2)
        print(order_df)
        
        

if __name__ == '__main__':
    main()