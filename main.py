import os
import tkinter as tk
from tkinter import filedialog
import yaml
import openpyxl

class YamlToSpreadsheetApp:
    def __init__(self, master):
        self.master = master
        master.title("YAML to Spreadsheet App")

        self.label = tk.Label(master, text="Select a directory containing YAML files:")
        self.label.pack()

        self.choose_dir_button = tk.Button(master, text="Choose Directory", command=self.choose_directory)
        self.choose_dir_button.pack()

        self.process_button = tk.Button(master, text="Process YAML files", command=self.process_yaml_files)
        self.process_button.pack()

    def choose_directory(self):
        self.directory_path = filedialog.askdirectory()

    def process_yaml_files(self):
        if not hasattr(self, 'directory_path'):
            return

        data = []
        for filename in os.listdir(self.directory_path):
            if filename.endswith(".yaml"):
                file_path = os.path.join(self.directory_path, filename)
                with open(file_path, 'r') as file:
                    yaml_data = yaml.safe_load(file)
                    data.append(yaml_data)

        self.create_spreadsheet(data)
        self.create_readme(data)

        tk.messagebox.showinfo("Success", "Process completed successfully!")

    def create_spreadsheet(self, data):
        wb = openpyxl.Workbook()
        ws = wb.active

        for row_num, yaml_data in enumerate(data, start=1):
            for col_num, (key, value) in enumerate(yaml_data.items(), start=1):
                ws.cell(row=row_num, column=col_num, value=f"{key}: {value}")

        spreadsheet_path = os.path.join(self.directory_path, 'output.xlsx')
        wb.save(spreadsheet_path)

    def create_readme(self, data):
        readme_path = os.path.join(self.directory_path, 'README.md')
        with open(readme_path, 'w') as readme_file:
            readme_file.write("# YAML Data\n\n")
            for yaml_data in data:
                for key, value in yaml_data.items():
                    readme_file.write(f"- **{key}**: {value}\n")
                readme_file.write("\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = YamlToSpreadsheetApp(root)
    root.mainloop()
