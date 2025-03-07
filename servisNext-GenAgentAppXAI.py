import tkinter as tk
import pandas as pd
import os

# Load data from Excel at the beginning of the code        
data1 = pd.read_excel(r"vec1.xlsx", header=None)
data2 = pd.read_excel(r"vec2.xlsx", header=None)
data3 = pd.read_excel(r"vec3.xlsx", header=None)
data4 = pd.read_excel(r"result.xlsx", header=None)

# Forming lists of values ​​for drop-down menus             
options1 = data1[0].tolist()
options2 = data2[0].tolist()
options3 = data3[0].tolist()

# Creating tables and vectors                             
column_names = data4.iloc[0, 1:].tolist()  
matrix_data = data4.iloc[1:, :].values  

# Function to save results to file                        
def save_to_file(filename, content):
    os.makedirs("D:\servisH", exist_ok=True)
    with open(os.path.join("D:\servisH", filename), "w") as file:
        file.write(content)

# Function to find matches in a table                     
def find_match():
    selected1 = dropdown1_var.get()
    selected2 = dropdown2_var.get()
    selected3 = dropdown3_var.get()

    # We obtain numerical values ​​corresponding to the selected categories             
    num1 = data1.loc[data1[0] == selected1, 1].values[0]
    num2 = data2.loc[data2[0] == selected2, 1].values[0]
    num3 = data3.loc[data3[0] == selected3, 1].values[0]

    saved_numbers = [num1, num2, num3]

    # Finding a matching column in a matrix                
    match_col = None
    for i in range(matrix_data.shape[1]):
        if list(matrix_data[:, i]) == saved_numbers:
            match_col = i
            break

    # Output of the result                                 
    if match_col is not None:
        result = column_names[match_col-1]  
    else:
        result = "Not enough information"

    result_label.config(text=result)
    save_to_file("matched_column.txt", result + "\n" + str(saved_numbers))

# Field Clearing Function                                 
def clear_fields():
    dropdown1_var.set(options1[0])
    dropdown2_var.set(options2[0])
    dropdown3_var.set(options3[0])
    result_label.config(text="")

# Creating a GUI (interface) Arial                        
root = tk.Tk()
root.title("Next-Gen Agent App Powered by Explainable AI")

tk.Label(root, text="temperature").grid(row=0, column=0)
tk.Label(root, text="blood pressure").grid(row=0, column=1)
tk.Label(root, text="external sign").grid(row=0, column=2)

dropdown1_var = tk.StringVar(value=options1[0])
dropdown2_var = tk.StringVar(value=options2[0])
dropdown3_var = tk.StringVar(value=options3[0])

tk.OptionMenu(root, dropdown1_var, *options1).grid(row=1, column=0)
tk.OptionMenu(root, dropdown2_var, *options2).grid(row=1, column=1)
tk.OptionMenu(root, dropdown3_var, *options3).grid(row=1, column=2)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.grid(row=2, column=0, columnspan=3)

tk.Button(root, text="Push", command=find_match).grid(row=3, column=0, columnspan=2)
tk.Button(root, text="Clean", command=clear_fields).grid(row=3, column=2)

root.mainloop()
