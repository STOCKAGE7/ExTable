import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

class PDFTableExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("ExTable")
        self.root.configure(bg='#2e2e2e')

        style = ttk.Style(self.root)
        default_font = ("Segoe UI", 10)
        self.root.option_add("*Font", default_font)
        style.theme_use("clam")

        style.configure("Treeview",
                        background="#2e2e2e",
                        fieldbackground="#2e2e2e",
                        foreground="white",
                        rowheight=30)
        style.configure("Treeview.Heading", background="#444", foreground="white")

        self.frame = tk.Frame(root, bg='#2e2e2e')
        self.frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        self.select_button = tk.Button(self.frame, text="Select PDF File", command=self.load_pdf,
                                       bg="#444", fg="white", relief=tk.FLAT)
        self.select_button.pack(pady=(0, 10))

        self.show_all_button = tk.Button(self.frame, text="Show All Tables", command=self.display_all_tables,
                                         bg="#555", fg="white", relief=tk.FLAT)
        self.show_all_button.pack(pady=(0, 10))

        self.export_button = tk.Button(self.frame, text="Export to Excel", command=self.export_to_excel,
                                       bg="#666", fg="white", relief=tk.FLAT)
        self.export_button.pack(pady=(0, 10))

        self.table_dropdown = ttk.Combobox(self.frame, state="readonly")
        self.table_dropdown.pack(pady=(0, 10))
        self.table_dropdown.bind("<<ComboboxSelected>>", self.display_table)

        self.canvas = tk.Canvas(self.frame, bg="#2e2e2e")
        self.scrollbar = tk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#2e2e2e")

        self.scrollable_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.tree = ttk.Treeview(self.frame, show="headings")
        self.tree.pack_forget()  # Hide initially
        self.tree.bind("<Double-1>", self.on_double_click)

        self.pdf_tables = []
        self.pdf_path = None

    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[["PDF Files", "*.pdf"]])
        if not file_path:
            return

        self.pdf_path = file_path
        self.pdf_tables.clear()
        try:
            with pdfplumber.open(file_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        if table:
                            self.pdf_tables.append((i + 1, j, table))
            if not self.pdf_tables:
                messagebox.showinfo("No Tables", "No tables found in the PDF.")
                return
            self.table_dropdown["values"] = [f"Page {p} - Table {t + 1}" for p, t, _ in self.pdf_tables]
            self.table_dropdown.current(0)
            self.display_table()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def display_table(self, event=None):
        self.clear_scrollable_frame()

        selected_index = self.table_dropdown.current()
        if selected_index == -1:
            return

        _, _, table_data = self.pdf_tables[selected_index]

        if not table_data:
            return

        headers = table_data[0]
        rows = table_data[1:]

        self.tree.pack_forget()
        tree = self.create_tree(self.scrollable_frame, headers, rows)
        tree.pack(fill=tk.X, pady=10)

    def display_all_tables(self):
        self.clear_scrollable_frame()
        self.tree.pack_forget()

        # Clear the dropdown selection when showing all tables
        self.table_dropdown.set('')  # This resets the dropdown to show no selection

        for page, table_index, table_data in self.pdf_tables:
            if not table_data:
                continue
            label = tk.Label(self.scrollable_frame, text=f"Page {page} - Table {table_index + 1}",
                             fg="white", bg="#2e2e2e", anchor="w")
            label.pack(fill=tk.X, pady=(10, 0))
            headers = table_data[0]
            rows = table_data[1:]
            tree = self.create_tree(self.scrollable_frame, headers, rows)
            tree.pack(fill=tk.X, pady=5)

    def create_tree(self, parent, headers, rows):
        tree = ttk.Treeview(parent, show="headings")
        tree["columns"] = list(range(len(headers)))

        for i, header in enumerate(headers):
            tree.heading(i, text=header if header else f"Col {i + 1}")
            tree.column(i, anchor="w", width=200, stretch=True)

        for row in rows:
            tree.insert("", "end", values=row)

        tree.bind("<Double-1>", self.on_double_click)
        return tree

    def clear_scrollable_frame(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

    def on_double_click(self, event):
        tree = event.widget
        item_id = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item_id or not column:
            return

        col_index = int(column.replace('#', '')) - 1
        item_values = tree.item(item_id, "values")

        if col_index >= len(item_values):
            return

        content = item_values[col_index]
        if not content:
            return

        top = tk.Toplevel(self.root)
        top.title("Full Cell Content")
        top.configure(bg='#2e2e2e')

        text = tk.Text(top, wrap="word", bg='#2e2e2e', fg="white", relief=tk.FLAT, padx=10, pady=10)
        text.insert("1.0", content)
        text.config(state="disabled")
        text.pack(fill=tk.BOTH, expand=True)

        top.geometry("500x300")

    def export_to_excel(self):
        if not self.pdf_tables:
            messagebox.showwarning("No Data", "No tables loaded to export.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[["Excel Files", "*.xlsx"]])
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                selected_index = self.table_dropdown.current()

                # If no table is selected (i.e., user clicked "Show All Tables"), export all tables
                if selected_index == -1:
                    for page, table_index, table_data in self.pdf_tables:
                        headers = table_data[0]
                        rows = table_data[1:]
                        df = pd.DataFrame(rows, columns=headers)
                        sheet_name = f"Page{page}_Table{table_index + 1}"
                        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                else:
                    # If a specific table is selected, export only that table
                    page, table_index, table_data = self.pdf_tables[selected_index]
                    headers = table_data[0]
                    rows = table_data[1:]
                    df = pd.DataFrame(rows, columns=headers)
                    sheet_name = f"Page{page}_Table{table_index + 1}"
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            messagebox.showinfo("Success", "Tables exported to Excel successfully!")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFTableExtractor(root)
    root.geometry("1000x700")
    root.mainloop()
