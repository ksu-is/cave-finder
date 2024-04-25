import arcpy
import pandas as pd

#path to workspace and lidar data

arcpy.env.workspace = r"D:/Desktop/ARCgistest.laz"  
lidar_data = "D:/Desktop/TEST.las"  

#convert lidar data into feature class

output_feature_class = "lidar_points"
arcpy.conversion.LASDatasetToFeatureClass(lidar_data, arcpy.env.workspace, output_feature_class)

#convert feature cass to dataframe

fields = ['SHAPE@X', 'SHAPE@Y', 'SHAPE@Z']
data = []
with arcpy.da.SearchCursor(output_feature_class, fields) as cursor:
    for row in cursor:
        data.append(row)
#convert data  to dataFrame
df = pd.DataFrame(data, columns=['X', 'Y', 'Z'])

# export to excel
excel_file = "lidar_data.xlsx"
df.to_excel(excel_file, index=False)

print("LIDAR data has been imported into Excel successfully.")
