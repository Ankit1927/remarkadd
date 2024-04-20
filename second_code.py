import pandas as pd

def add_remark(df):
    for index, row in df.iterrows():
        if "UPI" in row['Narration'] and pd.notnull(row['Deposit Amt.']) <= 999 :
            if pd.isnull(row['REMARK']):
                df.at[index, 'REMARK'] = "PERSONAL USE ONLINE DEBIT"
    return df

def main():
    # Prompt user for file path
    file_path = input("Enter the file path: ")

    # Read Excel file
    df = pd.read_excel(file_path)

    # Apply the operation
    df = add_remark(df)

    # Save output to a new Excel file
    output_file_path = "output.xlsx"
    df.to_excel(output_file_path, index=False)
    print("Output saved to", output_file_path)

if __name__ == "__main__":
    main()
