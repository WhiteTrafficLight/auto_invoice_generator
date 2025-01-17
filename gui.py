import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkinter.filedialog import askopenfilename, askdirectory
from invoice_generator import InvoiceGenerator

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Auto-Generator")

        # Initialize InvoiceGenerator with database support
        self.template_path = "templates/invoice_template.docx"
        self.output_dir = "output"
        self.generator = InvoiceGenerator(self.template_path, self.output_dir)

        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Customer Selection
        tk.Label(self.root, text="Customer:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.customer_var = tk.StringVar()
        self.customer_dropdown = ttk.Combobox(self.root, textvariable=self.customer_var, width=28)
        self.update_customer_dropdown()
        self.customer_dropdown.grid(row=0, column=1, padx=10, pady=5)

        # Button to Add New Customer
        ttk.Button(self.root, text="Add Customer", command=self.add_customer_popup).grid(row=0, column=2, padx=10, pady=5)

        # Date Selection
        tk.Label(self.root, text="Select Year:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.year_var = tk.StringVar()
        self.year_entry = ttk.Combobox(self.root, textvariable=self.year_var, width=28)
        self.year_entry['values'] = [str(year) for year in range(2000, 2100)]
        self.year_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(self.root, text="Select Month:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.month_var = tk.StringVar()
        self.month_entry = ttk.Combobox(self.root, textvariable=self.month_var, width=28)
        self.month_entry['values'] = [str(month).zfill(2) for month in range(1, 13)]
        self.month_entry.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(self.root, text="Select Day:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.day_var = tk.StringVar()
        self.day_entry = ttk.Combobox(self.root, textvariable=self.day_var, width=28)
        self.day_entry['values'] = [str(day).zfill(2) for day in range(1, 32)]
        self.day_entry.grid(row=3, column=1, padx=10, pady=5)

        # Quantity Inputs
        tk.Label(self.root, text="Quantity Item 1:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.quantity_item1_entry = ttk.Entry(self.root, width=30)
        self.quantity_item1_entry.grid(row=4, column=1, padx=10, pady=5)

        tk.Label(self.root, text="Quantity Item 2:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.quantity_item2_entry = ttk.Entry(self.root, width=30)
        self.quantity_item2_entry.grid(row=5, column=1, padx=10, pady=5)

        # Buttons
        ttk.Button(self.root, text="Generate Invoice", command=self.generate_invoice).grid(row=6, column=0, columnspan=3, pady=10)
        ttk.Button(self.root, text="Select Template", command=self.select_template).grid(row=7, column=0, padx=10, pady=5)
        ttk.Button(self.root, text="Set Output Directory", command=self.set_output_directory).grid(row=7, column=1, padx=10, pady=5)

    def update_customer_dropdown(self):
        """
        Update the customer dropdown with the latest list of customers from the database.
        """
        customers = self.generator.get_customers()
        self.customer_dropdown['values'] = customers

    def add_customer_popup(self):
        """
        Open a popup window to add a new customer.
        """
        customer_name = simpledialog.askstring("Add Customer", "Enter new customer name:")
        if customer_name:
            if self.generator.add_customer(customer_name):
                messagebox.showinfo("Success", f"Customer '{customer_name}' added.")
                self.update_customer_dropdown()
            else:
                messagebox.showerror("Error", f"Customer '{customer_name}' already exists.")

    def generate_invoice(self):
        # Gather user inputs
        customer = self.customer_var.get()
        year = self.year_var.get()
        month = self.month_var.get()
        day = self.day_var.get()
        quantity_item1 = self.quantity_item1_entry.get()
        quantity_item2 = self.quantity_item2_entry.get()

        # Validate inputs
        if not customer or not year or not month or not day or not quantity_item1 or not quantity_item2:
            messagebox.showerror("Input Error", "All fields are required!")
            return

        if not quantity_item1.isdigit() or not quantity_item2.isdigit():
            messagebox.showerror("Input Error", "Quantities must be numeric!")
            return

        # Prepare user data dictionary
        user_data = {
            "customer": customer,
            "year": year,
            "month": int(month),
            "day": int(day),
            "quantity_item1": quantity_item1,
            "quantity_item2": quantity_item2,
        }

        try:
            # Generate the invoice
            invoice_path = self.generator.generate_invoice(user_data)
            messagebox.showinfo("Success", f"Invoice generated at:\n{invoice_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def select_template(self):
        template_path = askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if template_path:
            self.template_path = template_path
            self.generator.template_path = template_path
            messagebox.showinfo("Template Selected", f"Template set to:\n{template_path}")

    def set_output_directory(self):
        output_dir = askdirectory()
        if output_dir:
            self.output_dir = output_dir
            self.generator.output_dir = output_dir
            messagebox.showinfo("Output Directory Set", f"Output directory set to:\n{output_dir}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()


