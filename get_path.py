import os
import openpyxl
current_directory = "D:\Fortis Export"
count = 0
# List all files in the directory
for root, dirs, files in os.walk(current_directory):
    for file in files:
        print(os.path.join(root, file))
        count+=1
print(count)