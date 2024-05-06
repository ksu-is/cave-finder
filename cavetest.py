import laspy
import pandas as pd

#Path to the LAZ file
input_laz_file = r"D:/Desktop/lidartest.laz"

#Path to the output Excel file
output_excel_file = r"D:/Desktop/lidarexcel.xlsx"

def read_laz(laz_file, excel_file, max_entries=3000):
    try:
        #read LAZ file
        las_file = laspy.read(laz_file)

        #Extract X, Y, Z coordinates and convert to lists
        x = list(las_file.x)[:max_entries]
        y = list(las_file.y)[:max_entries]
        z = list(las_file.z)[:max_entries]

        #Create DataFrame
        df = pd.DataFrame({"X": x, "Y": y, "Z": z})

        #Export DataFrame to Excel
        df.to_excel(excel_file, index=False)
        print(f"Point data exported to {excel_file}")

    except Exception as e:
        print("An error occurred:", str(e))

try:
    read_laz(input_laz_file, output_excel_file)
except Exception as e:
    print("An error occurred:", str(e))
