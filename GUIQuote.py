import sys
from fpdf import FPDF
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    try:
        base_path = sys._MEIPASS  # When running from .exe
    except AttributeError:
        base_path = os.path.abspath(".")  # When running as script
    return os.path.join(base_path, relative_path)


class QuotationPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.company_name = "DRAPES AND FITOUT FZE"
        self.company_address = "OFFICE B40-003, BLOCK B, SHARJAH, UAE"
        self.company_phone = "+971 58 580 9365"
        self.company_vat = "104706477700003"
        self.logo_path = resource_path("Drapes And Fiouts.jpeg")
        self.vat_rate = 5  # 5% VAT
        self.quote_date = datetime.now().strftime("%d-%b-%Y")
        self.quote_number = "QT-" + datetime.now().strftime("%Y%m%d-%H%M")
        self.terms_and_conditions = [
            "1. This quotation is valid for 30 days from the date of issue.",
            "2. Entry permits should be arranged by the client.",
            "3. Orders once confirmed cannot be cancelled.",
            "4. Rectifying ceiling issues does not fall under our scope of work.",
            "5. Ceiling issues will be flagged at time of measurement.",
            "6. Space must be ready before installation.",
            "7. Deep cleaning must be done before installation.",
            "8. Payment terms: 80% advance, 20% upon installation.",
            "9. Cheques payable to: DRAPES AND FITOUTS TRADING FZE",
            "10. Bank Details:\n   - Bank: RAK Bank\n   - Account: 0353425596001\n   - IBAN: AE31 0400 0003 5342 5596 001"
        ]

           # [Rest of your QuotationPDF methods remain unchanged...]

    def header(self):
        # Add logo if exists
        if os.path.exists(self.logo_path):
            self.image(self.logo_path, 10, 8, 33)

        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)  # Dark blue color
        self.cell(0, 10, self.company_name, 0, 1, 'C')

        self.set_font('Arial', '', 10)
        self.cell(0, 6, self.company_address, 0, 1, 'C')
        self.cell(0, 6, f"Phone: {self.company_phone} | VAT: {self.company_vat}", 0, 1, 'C')
        self.cell(0, 6, f"Quotation No: {self.quote_number} | Date: {self.quote_date}", 0, 1, 'C')
        self.ln(10)

    def client_details(self, name, phone, address, project):
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(200, 200, 200)  # Light gray background
        self.cell(0, 10, 'CLIENT DETAILS', 0, 1, 'L', fill=True)

        self.set_font('Arial', '', 10)
        self.cell(40, 6, 'Client Name:', 0, 0)
        self.cell(0, 6, name, 0, 1)
        self.cell(40, 6, 'Phone Number:', 0, 0)
        self.cell(0, 6, phone, 0, 1)
        self.cell(40, 6, 'Address:', 0, 0)
        self.multi_cell(0, 6, address, 0, 1)
        self.cell(40, 6, 'Project Name:', 0, 0)
        self.cell(0, 6, project, 0, 1)
        self.ln(10)

    def add_item(self, room, description, quantity, unit_price, total_price):
        self.set_font('Arial', '', 10)
        self.cell(40, 8, room, 1)
        self.cell(80, 8, description, 1)
        self.cell(20, 8, str(quantity), 1, 0, 'C')
        self.cell(25, 8, f"{unit_price:,.2f} AED", 1, 0, 'R')
        self.cell(25, 8, f"{total_price:,.2f} AED", 1, 1, 'R')

    def add_total(self, total):
        self.ln(8)
        self.set_font('Arial', 'B', 10)
        vat_amount = total * (self.vat_rate / 100)
        grand_total = total + vat_amount

        self.cell(140, 8, 'Subtotal:', 1, 0, 'R')
        self.cell(50, 8, f"{total:,.2f} AED", 1, 1, 'R')

        self.cell(140, 8, f'VAT ({self.vat_rate}%):', 1, 0, 'R')
        self.cell(50, 8, f"{vat_amount:,.2f} AED", 1, 1, 'R')

        self.set_font('Arial', 'B', 12)
        self.cell(140, 10, 'GRAND TOTAL:', 1, 0, 'R')
        self.cell(50, 10, f"{grand_total:,.2f} AED", 1, 1, 'R')

    def add_terms_and_conditions(self):
        self.ln(10)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'TERMS AND CONDITIONS', 0, 1)

        self.set_font('Arial', '', 10)
        for term in self.terms_and_conditions:
            self.multi_cell(0, 5, term, 0, 1)

        self.ln(5)
        self.set_font('Arial', 'I', 10)
        self.cell(0, 10, "Thank you for your business!", 0, 1, 'C')


# [Keep all your existing QuotationApp class code exactly as is]

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drapes and Fitout Quotation Generator")
        self.root.geometry("900x700")

        # Configure styles
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Accent.TButton', foreground='white', background='#4CAF50', font=('Arial', 10, 'bold'))

        # Create main container with scrollbar
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Now use scrollable_frame instead of main_frame for all content
        self.main_frame = self.scrollable_frame

        # Client Information Section
        self.create_client_section()
        # Product Entry Section
        self.create_product_section()
        # Items Display Section
        self.create_items_display()
        # Buttons Section
        self.create_buttons_section()

        # Initialize data storage
        self.items = []
        self.current_product_type = None

        # Add mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling for the canvas"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ... [keep all your other existing methods exactly the same] ...

    # ... [keep all other methods exactly the same until create_buttons_section] ...

    def create_buttons_section(self):
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        # Modified generate button with new style
        generate_btn = ttk.Button(
            button_frame,
            text="Generate Quotation PDF",
            command=self.generate_quotation,
            style='Accent.TButton'
        )
        generate_btn.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=5)

        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

        # Initialize data storage
        self.items = []
        self.current_product_type = None

    def create_client_section(self):
        client_frame = ttk.LabelFrame(self.main_frame, text="Client Information")
        client_frame.pack(fill=tk.X, pady=5)

        fields = [
            ("Client Name:", "client_name"),
            ("Phone Number:", "client_phone"),
            ("Address:", "client_address"),
            ("Project Name:", "project_name")
        ]

        for i, (label, var_name) in enumerate(fields):
            setattr(self, var_name, tk.StringVar())
            ttk.Label(client_frame, text=label).grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Entry(client_frame, textvariable=getattr(self, var_name), width=40).grid(
                row=i, column=1, sticky=tk.W, padx=5, pady=2)

    def create_product_section(self):
        product_frame = ttk.LabelFrame(self.main_frame, text="Product Details")
        product_frame.pack(fill=tk.X, pady=5)

        # Product Type Selection
        ttk.Label(product_frame, text="Product Type:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.product_type = tk.StringVar()
        ttk.Radiobutton(product_frame, text="Blind", variable=self.product_type,
                        value="blind", command=self.toggle_product_fields).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(product_frame, text="Curtain", variable=self.product_type,
                        value="curtain", command=self.toggle_product_fields).grid(row=0, column=2, sticky=tk.W)

        # Common Fields
        ttk.Label(product_frame, text="Room Name:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.room_name = tk.StringVar()
        ttk.Entry(product_frame, textvariable=self.room_name).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(product_frame, text="Width (m):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.width = tk.DoubleVar()
        ttk.Entry(product_frame, textvariable=self.width).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(product_frame, text="Height (m):").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.height = tk.DoubleVar()
        ttk.Entry(product_frame, textvariable=self.height).grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)

        # Motorization
        self.motorized = tk.BooleanVar()
        ttk.Checkbutton(product_frame, text="Motorized", variable=self.motorized,
                        command=self.toggle_motor_fields).grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)

        # Blind-specific fields (hidden initially)
        self.blind_frame = ttk.Frame(product_frame)
        ttk.Label(self.blind_frame, text="Price per sqm:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.blind_price = tk.DoubleVar()
        ttk.Entry(self.blind_frame, textvariable=self.blind_price).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        # Curtain-specific fields (hidden initially)
        self.curtain_frame = ttk.Frame(product_frame)
        fields = [
            ("Fullness Factor:", "fullness_factor", 2.0),
            ("Fabric Width (m):", "fabric_width", 2.8),
            ("Stitching Cost/m:", "stitching_cost", 40),
            ("Fabric Price/m:", "fabric_price", 60),
            ("Track Price/m:", "track_price", 40)
        ]

        for i, (label, var_name, default) in enumerate(fields):
            setattr(self, var_name, tk.DoubleVar(value=default))
            ttk.Label(self.curtain_frame, text=label).grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            ttk.Entry(self.curtain_frame, textvariable=getattr(self, var_name)).grid(
                row=i, column=1, sticky=tk.W, padx=5, pady=2)

        # Motor fields (hidden initially)
        self.motor_frame = ttk.Frame(product_frame)
        ttk.Label(self.motor_frame, text="Motor Price:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.motor_price = tk.DoubleVar(value=300)
        ttk.Entry(self.motor_frame, textvariable=self.motor_price).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(self.motor_frame, text="Remote Price:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.remote_price = tk.DoubleVar(value=150)
        ttk.Entry(self.motor_frame, textvariable=self.remote_price).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        # Add Item Button
        ttk.Button(product_frame, text="Add Item", command=self.add_item).grid(
            row=20, column=0, columnspan=3, pady=10)

    def toggle_product_fields(self):
        if self.product_type.get() == "blind":
            self.curtain_frame.grid_forget()
            self.blind_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)
        else:
            self.blind_frame.grid_forget()
            self.curtain_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)

    def toggle_motor_fields(self):
        if self.motorized.get():
            self.motor_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)
        else:
            self.motor_frame.grid_forget()

    def create_items_display(self):
        items_frame = ttk.LabelFrame(self.main_frame, text="Items in Quotation")
        items_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Create Treeview
        self.tree = ttk.Treeview(items_frame, columns=("Room", "Description", "Qty", "Price"), show="headings")
        self.tree.heading("Room", text="Room")
        self.tree.heading("Description", text="Description")
        self.tree.heading("Qty", text="Qty")
        self.tree.heading("Price", text="Price (AED)")

        self.tree.column("Room", width=100)
        self.tree.column("Description", width=400)
        self.tree.column("Qty", width=50, anchor=tk.CENTER)
        self.tree.column("Price", width=100, anchor=tk.E)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Delete button
        ttk.Button(items_frame, text="Delete Selected", command=self.delete_item).pack(side=tk.LEFT, pady=5)

    def create_buttons_section(self):
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="Generate Quotation", command=self.generate_quotation).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

    def add_item(self):
        try:
            room = self.room_name.get()
            if not room:
                messagebox.showerror("Error", "Please enter a room name")
                return

            width = self.width.get()
            height = self.height.get()

            if self.product_type.get() == "blind":
                price_sqm = self.blind_price.get()
                area = width * height
                total_price = area * price_sqm + 100  # Base installation

                description = f"{'Motorized' if self.motorized.get() else 'Manual'} Blind ({width}m × {height}m)"

                if self.motorized.get():
                    motor_price = self.motor_price.get()
                    remote_price = self.remote_price.get()
                    total_price += motor_price + remote_price + 100  # Motor installation

            else:  # Curtain
                fullness = self.fullness_factor.get()
                fabric_width = self.fabric_width.get()
                fabric_req = (width * fullness) / fabric_width
                length_req = height + 0.3

                fabric_cost = fabric_req * length_req * self.fabric_price.get()
                stitching_cost = fabric_req * length_req * self.stitching_cost.get()
                track_cost = width * self.track_price.get()
                total_price = fabric_cost + stitching_cost + track_cost + 100

                description = f"{'Motorized' if self.motorized.get() else 'Manual'} Curtain (Fabric: {length_req:.2f}m × {fabric_req:.2f} widths)"

                if self.motorized.get():
                    total_price += self.motor_price.get() + self.remote_price.get() + 100

            # Add to items list and treeview
            item = {
                'room_name': room,
                'description': description,
                'quantity': 1,
                'unit_price': round(total_price, 2),
                'total_price': round(total_price, 2)
            }

            self.items.append(item)
            self.tree.insert("", tk.END, values=(
                item['room_name'],
                item['description'],
                item['quantity'],
                f"{item['total_price']:,.2f}"
            ))

            # Clear fields for next entry
            self.room_name.set("")
            self.width.set(0)
            self.height.set(0)

        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")

    def delete_item(self):
        selected = self.tree.selection()
        if not selected:
            return

        for item in selected:
            index = self.tree.index(item)
            self.tree.delete(item)
            del self.items[index]

    def clear_all(self):
        if messagebox.askyesno("Confirm", "Clear all items and client information?"):
            self.items.clear()
            self.tree.delete(*self.tree.get_children())

            # Clear all input fields
            for var in [self.client_name, self.client_phone, self.client_address,
                        self.project_name, self.room_name, self.width, self.height]:
                var.set("")

            self.motorized.set(False)
            self.product_type.set("")

    def generate_quotation(self):
        if not self.items:
            messagebox.showerror("Error", "No items added to quotation")
            return

        client_info = (
            self.client_name.get(),
            self.client_phone.get(),
            self.client_address.get(),
            self.project_name.get()
        )

        if not all(client_info):
            messagebox.showerror("Error", "Please complete all client information")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Save Quotation As"
        )

        if not file_path:
            return

        pdf = QuotationPDF()
        pdf.add_page()
        pdf.client_details(*client_info)

        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 10, 'Room Name', 1)
        pdf.cell(80, 10, 'Description', 1)
        pdf.cell(20, 10, 'Quantity', 1)
        pdf.cell(25, 10, 'Unit Price', 1)
        pdf.cell(25, 10, 'Total Price', 1, 1)

        total_cost = sum(item['total_price'] for item in self.items)

        for item in self.items:
            pdf.add_item(item['room_name'], item['description'],
                         item['quantity'], item['unit_price'], item['total_price'])

        pdf.add_total(total_cost)
        pdf.add_terms_and_conditions()

        try:
            pdf.output(file_path)
            messagebox.showinfo("Success", f"Quotation saved successfully:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationApp(root)
    root.mainloop()