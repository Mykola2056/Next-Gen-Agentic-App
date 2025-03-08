import tkinter as tk
import pandas as pd
import os

def load_data(file_paths):
    """Loads data from Excel files."""
    data = {}
    for name, path in file_paths.items():
        try:
            data[name] = pd.read_excel(path, header=None)
        except FileNotFoundError:
            print(f"File not found: {path}")
            return None
    return data

def save_to_file(filename, content):
    """Saves content to a file."""
    os.makedirs(r"D:\servisH", exist_ok=True)
    with open(os.path.join(r"D:\servisH", filename), "w") as file:
        file.write(content)

def find_match():
    """Finds matches in a table based on dropdown selections."""
    selected1 = dropdown1_var.get()
    selected2 = dropdown2_var.get()
    selected3 = dropdown3_var.get()

    num1 = data["data1"].loc[data["data1"][0] == selected1, 1].values[0]
    num2 = data["data2"].loc[data["data2"][0] == selected2, 1].values[0]
    num3 = data["data3"].loc[data["data3"][0] == selected3, 1].values[0]

    saved_numbers = [num1, num2, num3]

    match_col = None
    for i in range(matrix_data.shape[1]):
        if list(matrix_data[:, i]) == saved_numbers:
            match_col = i
            break

    if match_col is not None:
        result = column_names[match_col - 1]
    else:
        result = "Not enough information"

    result_label.config(text=result)
    save_to_file("matched_column.txt", result + "\n" + str(saved_numbers))

def clear_fields():
    """Clears the dropdown selections and result label."""
    dropdown1_var.set(options1[0])
    dropdown2_var.set(options2[0])
    dropdown3_var.set(options3[0])
    result_label.config(text="")

# Load data from Excel files
file_paths = {
    "data1": r"vec1.xlsx",
    "data2": r"vec2.xlsx",
    "data3": r"vec3.xlsx",
    "data4": r"result.xlsx",
}

data = load_data(file_paths)

if data is None:
    exit() # Exit if data loading failed.

options1 = data["data1"][0].tolist()
options2 = data["data2"][0].tolist()
options3 = data["data3"][0].tolist()

column_names = data["data4"].iloc[0, 1:].tolist()
matrix_data = data["data4"].iloc[1:, :].values

# GUI setup
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
