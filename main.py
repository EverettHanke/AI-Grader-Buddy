import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

class ExcelApp:
    #*********************************************
    # INITIALIZES THE APPLICATION
    #*********************************************
    def __init__(self, root):
        self.root = root
        self.root.title("AI Grader Buddy")

        # FRAME FOR BUTTONS
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=10)
        #UPLOAD BUTTON
        self.upload_button = tk.Button(self.frame, text="Upload Excel File", command=self.upload_file)
        self.upload_button.pack(side=tk.LEFT, padx=5)

        self.label = tk.Label(self.frame, text="No file selected", fg="gray")
        self.label.pack(side=tk.LEFT, padx=5)

        # DISPLAYS THE EXCEL DATA UTILIZING TREEVIEW
        self.tree = ttk.Treeview(self.root, columns=[], show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ADDS A SCROLLBAR TO TREEVIEW
        self.scroll_x = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.scroll_y = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)

    #*********************************************
    # UPLOADS THE EXCEL FILE
    #*********************************************
    def upload_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        try:
            # Read the Excel file using pandas
            data = pd.read_excel(file_path)
            self.display_data(data)
            self.label.config(text=f"Loaded: {file_path.split('/')[-1]}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    #*********************************************
    # DISPLAYS THE DATA IN THE TREEVIEW
    #*********************************************
    def display_data(self, data):
        # Clear previous data
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = list(data.columns)

        # Set up columns and headings
        for col in data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)

        # Insert data into Treeview
        for _, row in data.iterrows():
            self.tree.insert("", tk.END, values=row.tolist())



#*********************************************
#RUNS THE APPLICATION
#*********************************************
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.geometry("800x600")
    root.mainloop()
