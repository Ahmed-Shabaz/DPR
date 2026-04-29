import pandas as pd
import os

# def extract_specific_cells(input_excel, output_excel, cells=["E16", "A17", "B11", "K6", "F8", "E11", "A12", "F5", "H18", "J2", "K8", "K2", "L6", "M9", "L5", "N17", "L8", "M12", "M11", "N18", "E14", "A15"]):
# def extract_specific_cells(input_excel, output_excel, 
#     cells=["E16", "A17", "B11", "K6", "F8", "E11", "A12", "F5", 
#            "H18", "J2", "K8", "K2", "L6", "M9", "L5", "N17", "L8", 
#            "M12", "M11", "N18", "E14", "A15", "C20", "D25", "F30",
#            "G11", "P12", "O8", "E7", "B6", "F7", "M8", "C20", "D25", "F30",
#            "G11", "P12", "O8", "E7", "B6", "F7", "M8", "M9", "Q16", "K15", "K7", "F6", "F10"
# N16", "H15", "N6", "F11", "O12", "E5", "I2", "B9", "L2", "l2", "D6", "J7"
# O12", "P16", P16", "O12", "S16", "R12", "P13", "I18", "M2", "H5", "G15", "I19", "Q20",

        #    "J16", "S17", "R13", "Q17", "H16", "O13", "L9", "F8", "Q20", "K5", "R16", "Q12",
        #    "R14", "J13", "Q10", "O19", "N15", "G18]):
        # M9", "L5", "E8", "J15", "J17", "I6", "R18", "L10", "F9", "K6", "G15",
        #    "G5", "J17", "R18", "E16", "A17", "B11", "D5", "H2", "L9"]):S18", "Q13", "R19", "I16", "R17
        # P19", "O15", "E17", "A18", "B12", "O16", "N12", "I15", "E45","A46", "B37",


def extract_specific_cells(input_excel, output_excel, 
    cells=["I17", "Q18", "P14"]):  
    results = []

    # Open Excel file
    xls = pd.ExcelFile(input_excel)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_excel, sheet_name=sheet_name, header=None)

        # Extract required cells
        extracted = {}
        extracted["File Name"] = sheet_name  # ✅ use sheet name instead of file name

        for cell in cells:
            col = ord(cell[0].upper()) - 65  # Convert column letter to index (A=0, B=1, etc.)
            row = int(cell[1:]) - 1          # Convert row number to 0-based index

            try:
                extracted[cell] = df.iat[row, col]
            except Exception:
                extracted[cell] = None  # If cell is out of range, set empty

        results.append(extracted)

    # Save results to new Excel file
    output_df = pd.DataFrame(results)
    output_df.to_excel(output_excel, index=False)

    print(f"✅ Extracted cells saved to: {output_excel}")


# Usage example
input_excel = r"Combined_Output_835.xlsx"
output_excel = r"Book1111.xlsx"

extract_specific_cells(input_excel, output_excel)
