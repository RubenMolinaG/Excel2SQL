from openpyxl import load_workbook

def main() -> bool:
    EXCEL_FILE_NAME = "Data.xlsx"
    CSV_FILE_NAME = "Data.csv"
    SQL_FILE_NAME = "KYC_INSERT_QUERY.sql"
    START_COL = 2

    if get_csv_from_excel(EXCEL_FILE_NAME, CSV_FILE_NAME, START_COL):
        if(set_sql_insert_query(SQL_FILE_NAME, read_csv_file(CSV_FILE_NAME))):
            return True    
    return False


def get_csv_from_excel(excel_file_name: str, csv_file_name: str, start_col: int) -> bool:
    wb = load_workbook(filename=f"{excel_file_name}")
    ws = wb.active
    try:
        with open(f"{csv_file_name}", 'w') as text_file:
            for row in ws.iter_rows(min_row=start_col):
                contador = 0
                for cell in row:
                    text_file.write(cell.value + ",") if contador < 3 else text_file.write(cell.value + "\n")
                    contador += 1
    except Exception as e:
        print(f"Error: {e}")
        return False
    return True

def read_csv_file(csv_file_name: str) -> str:
    output: str = ""
    try:
        with open(f"{csv_file_name}", "r") as csv_file:
            for row in csv_file:
                elements = row.split(",")
                col01 = elements[0]
                col02 = elements[1]
                col03 = elements[2]
                col04 = elements[3].split("\n")[0]
                output += f"INSERT INTO(COL01, COL02, COL03, COL04) VALUES({col01}, {col02}, {col03}, {col04}); \n"
    except Exception as e:
        print(f"Error: {e}")
    return output
            
def set_sql_insert_query(sql_file_name: str, text_query: str) -> bool:
    queries: list[str] = text_query.split("\n")
    try:
        with open(f"{sql_file_name}", "w") as sql_file:
            sql_file.write("BEGIN BLOCK\n")
            for query in queries:
                if query != "":
                    sql_file.write(query + "\n")
            sql_file.write("END")
    except Exception as e:
        print(f"Error: {e}")
        return False
    return True

if __name__ == '__main__':
    print("SQL file created succesfully.") if main() else print("An error ocurred while creating the SQL file.")
    