import pandas as pd
import logging

def update_remarks(excel_file):
    """
    Update remarks in the Excel file based on specific conditions.

    Parameters:
    excel_file (str): The path to the Excel file.

    Returns:
    str: The path to the updated Excel file.
    """
    # Configure logging
    logging.info("Reading Excel file...")
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        # Handle file not found error
        logging.error("File not found.")
        return None
    
    # Calculate progress update intervals
    total_rows = len(df)
    progress_step = total_rows // 10  # Update progress every 10%
    current_progress = 0
    
    # Update remarks based on specific conditions
    logging.info("Updating remarks...")
    for index, row in df.iterrows():
        # Extract relevant data from the row
        narration = row.get('Narration', '')
        withdrawal_amt = row.get('Withdrawal Amt.', None)
        deposit_amt = row.get('Deposit Amt.', None)

        # Check conditions and update remarks accordingly
        if "POS 403875XXXXXX4387 UPGOVTOTHDRCARD" in narration and withdrawal_amt is not None:
            df.at[index, 'REMARK'] = "CHALAN PAYMENT ONLINE DEBIT"
        elif "UPI-SBIMOPS-SBIMOPS@SBI-SBIN0016209" in narration:
            df.at[index, 'REMARK'] = "CHALAN PAYMENT ONLINE DEBIT"
        elif "IMPS" in narration:
            if withdrawal_amt is not None and withdrawal_amt >= 5000:
                df.at[index, 'REMARK'] = "PERSONAL USE ONLINE DEBIT"
            elif withdrawal_amt is None and deposit_amt is not None:
                df.at[index, 'REMARK'] = "CHALAN PAYMENT ONLINE CREDIT"
        elif "UPI-" in narration and deposit_amt is not None and deposit_amt >= 2000:
            df.at[index, 'REMARK'] = "CHALAN PAYMENT ONLINE CREDIT"
        elif ".IMPS" in narration:
            df.at[index, 'REMARK'] = "BANK CHARGES DEBIT"
        elif "CASH DEPOSIT" in narration:
            df.at[index, 'REMARK'] = "CASH DEPOSIT"
        
        # Update progress
        if index >= current_progress:
            logging.info(f"Progress: {round((index / total_rows) * 100)}%")
            current_progress += progress_step

    # Save the updated DataFrame to a new Excel file
    updated_excel_file = excel_file.split('.')[0] + "_updated.xlsx"
    df.to_excel(updated_excel_file, index=False)
    logging.info("Remarks updated successfully.")
    return updated_excel_file

def add_remark(df):
    """
    Add remarks to the DataFrame based on specific conditions.

    Parameters:
    df (DataFrame): The DataFrame containing transaction data.

    Returns:
    DataFrame: The DataFrame with remarks added.
    """
    for index, row in df.iterrows():
        # Extract relevant data from the row
        narration = row.get('Narration', '')
        deposit_amt = row.get('Deposit Amt.', None)

        # Check condition and add remark accordingly
        if "UPI" in narration and deposit_amt is not None and deposit_amt <= 999:
            if pd.isnull(row['REMARK']):
                df.at[index, 'REMARK'] = "PERSONAL USE ONLINE CREDIT"
    return df

def main():
    """
    Main function to execute the script.
    """
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # Prompt user for file path
    excel_file = input("Enter the path of the Excel file: ")

    # Update remarks in the Excel file
    updated_file = update_remarks(excel_file)

    if updated_file:
        # Read the updated Excel file
        df_updated = pd.read_excel(updated_file)

        # Apply additional remarks
        df_updated = add_remark(df_updated)

        # Save output to a new Excel file
        output_file_path = "output.xlsx"
        df_updated.to_excel(output_file_path, index=False)
        logging.info("Output saved to %s", output_file_path)

if __name__ == "__main__":
    main()
