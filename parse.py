if __name__ == '__main__':

    import openpyxl

    # Define variable to load the dataframe
    dataframe = openpyxl.load_workbook("./files/test2.xlsx")

    # Define variable to read sheet
    dataframe1 = dataframe.active

    # Iterate the loop to read the cell values
    for row in dataframe1.iter_rows(1, dataframe1.max_row):
      for col in range(0, 3):
        if col == 2:
          print(row[col].value)
        else:
          print(row[col].value, end="|")