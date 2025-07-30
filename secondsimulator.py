import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd
import io
from PIL import Image
import openpyxl

# --- NEW IMPORTS FOR PDF GENERATION ---
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
# For Hebrew support in ReportLab:
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
from reportlab.lib.styles import ParagraphStyle

# Register a font that supports Hebrew characters (e.g., Arial Unicode MS or DejaVuSans)
# You might need to provide the full path to a .ttf file if it's not in your system's font paths
# For example, download 'DejaVuSans.ttf' and place it in your script's directory, or
# 'arial.ttf' if you are on Windows and it's typically found at C:/Windows/Fonts/arial.ttf

# --- MODIFIED FONT REGISTRATION ---
try:
    # Register the regular font
    pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))
    # Register the bold font (assuming DejaVuSans-Bold.ttf exists)
    pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', 'DejaVuSans-Bold.ttf'))
    # Register the font family to link regular and bold versions
    pdfmetrics.registerFontFamily('DejaVuSans',
                                  normal='DejaVuSans',
                                  bold='DejaVuSans-Bold',
                                  italic='DejaVuSans', # If you don't have italic, map to normal
                                  boldItalic='DejaVuSans-Bold') # If you don't have bold italic, map to bold

    # Define a style for Hebrew text that aligns right-to-left
    # Note: ReportLab's bidi support can be complex. For simple strings, it often works.
    # For complex RTL layouts, more advanced methods or external libraries might be needed.
    # We will set alignment to TA_RIGHT for Hebrew.
    heb_style = ParagraphStyle(name='Hebrew', fontName='DejaVuSans', fontSize=10, alignment=TA_RIGHT)
    # Use the 'DejaVuSans-Bold' font name directly for headings if you prefer,
    # or rely on the <b> tag and registerFontFamily for automatic bolding.
    # For headings, it's often better to explicitly set the fontName to the bold variant
    # if you want consistent bolding without relying on HTML tags inside Paragraphs.
    heb_heading_style = ParagraphStyle(name='HebrewHeading', fontName='DejaVuSans-Bold', fontSize=14, alignment=TA_RIGHT, spaceAfter=6)
    heb_subheading_style = ParagraphStyle(name='HebrewSubHeading', fontName='DejaVuSans-Bold', fontSize=12, alignment=TA_RIGHT, spaceAfter=4)
except Exception as e:
    messagebox.showwarning("Font Warning", f"Could not load DejaVuSans font for PDF. Hebrew text may not display correctly: {e}\nMake sure 'DejaVuSans.ttf' and 'DejaVuSans-Bold.ttf' are in the script's directory or provide full paths.")
    # Fallback to a default font if DejaVuSans isn't found
    heb_style = ParagraphStyle(name='Hebrew', fontName='Helvetica', fontSize=10, alignment=TA_RIGHT)
    heb_heading_style = ParagraphStyle(name='HebrewHeading', fontName='Helvetica-Bold', fontSize=14, alignment=TA_RIGHT, spaceAfter=6)
    heb_subheading_style = ParagraphStyle(name='HebrewSubHeading', fontName='Helvetica-Bold', fontSize=12, alignment=TA_RIGHT, spaceAfter=4)
    
# Constants for fees
LAWYER_FEE_RATE = 0.01
BROKER_FEE_RATE = 0.02

def calculate_purchase_tax(price):
    # Tax brackets and rates for purchase tax (assuming Israeli tax law for example)
    # These are illustrative and should be updated with actual current rates if this is for real use.
    brackets = [
        (0, 6055070, 8),
        (6055070, float('inf'), 10),
    ]
    tax = 0
    remaining_price = price
    for low, high, rate in brackets:
        if remaining_price > low:
            taxable_in_current_bracket = min(high, remaining_price) - low if remaining_price > low else 0
            tax += taxable_in_current_bracket * rate / 100
            remaining_price -= taxable_in_current_bracket
            if remaining_price <= 0:
                break
    return tax

def estimate_lawyer_fee(price):
    return price * LAWYER_FEE_RATE

def estimate_broker_fee(price):
    return price * BROKER_FEE_RATE

def calculate_monthly_payment(loan_amount, annual_rate, years):
    if loan_amount <= 0 or years <= 0:
        return 0.0
    months = years * 12
    monthly_rate = (annual_rate / 100) / 12
    if monthly_rate == 0:
        return loan_amount / months if months > 0 else 0.0
    
    if abs(monthly_rate) < 1e-9:
        return loan_amount / months if months > 0 else 0.0

    try:
        power_term = (1 + monthly_rate) ** months
    except OverflowError:
        return float('inf') 

    denominator = (power_term - 1)
    if denominator == 0: 
        return loan_amount / months if months > 0 else 0.0
    
    return loan_amount * (monthly_rate * power_term) / denominator

def generate_amortization_df(loan_amount, annual_rate, years):
    if loan_amount <= 0 or annual_rate < 0 or years <= 0:
        return pd.DataFrame() 

    current_monthly_payment = calculate_monthly_payment(loan_amount, annual_rate, years)
    months = years * 12
    balance = loan_amount
    data = []

    for month in range(1, months + 1):
        if balance <= 0:
            break

        interest = balance * (annual_rate / 100) / 12
        principal = current_monthly_payment - interest
        
        principal = min(principal, balance)
        
        balance -= principal
        
        data.append({
            "חודש": round(month, 2), 
            "קרן": round(principal, 2), 
            "ריבית": round(interest, 2),
            "יתרה": round(max(balance, 0), 2),
            "תשלום חודשי": round(current_monthly_payment, 2)
        })

    return pd.DataFrame(data)

# --- NEW FUNCTION FOR ERROR MESSAGES WITH COPY ---
def show_error_with_copy(title, message, parent=None):
    top = tk.Toplevel(parent)
    top.title(title)
    top.transient(parent)
    top.grab_set()
    top.lift() # Bring to front

    # Center the Toplevel window
    if parent:
        parent.update_idletasks() # Ensure parent geometry is up-to-date
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (top.winfo_reqwidth() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (top.winfo_reqheight() // 2)
        top.geometry(f"+{x}+{y}")

    tk.Label(top, text=message, wraplength=400, justify="center", padx=10, pady=10).pack()

    button_frame = ttk.Frame(top)
    button_frame.pack(pady=5)

    def copy_to_clipboard():
        top.clipboard_clear()
        top.clipboard_append(message)
        top.update() # Now it stays on the clipboard after the window is closed
        top.destroy()

    def close_window():
        top.destroy()

    ttk.Button(button_frame, text="סגור", command=close_window).pack(side="left", padx=5)
    ttk.Button(button_frame, text="העתק שגיאה", command=copy_to_clipboard).pack(side="right", padx=5)

    top.wait_window(top)


class PropertyTab:
    def __init__(self, parent, idx, root_window):
        self.root = root_window
        self.idx = idx
        self.frame = ttk.Frame(parent)
        self.frame.pack(expand=True, fill="both")

        self.canvas = tk.Canvas(self.frame, borderwidth=0, background="#f0f0f0")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.bind("<Button-4>", self._on_mousewheel_up)
        self.canvas.bind("<Button-5>", self._on_mousewheel_down)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        self.content_frame = ttk.Frame(self.canvas, padding="10 10 10 10")
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        self.input_frame = ttk.Frame(self.content_frame) 
        self.input_frame.pack(pady=10) 

        def new_entry():
            return tk.Entry(self.input_frame, justify='right', width=15, font=("Arial", 11))

        padx = 5
        pady = 3
        r = 0 

        ttk.Label(self.input_frame, text="Alias:").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.alias_entry = new_entry()
        self.alias_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="Link:").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.link_entry = new_entry()
        self.link_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="מחיר דירה (₪):").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.price_entry = new_entry()
        self.price_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="מטר מרובע (שטח):").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.area_entry = new_entry()
        self.area_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="אחוז מימון (LTV) %:").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.ltv_entry = new_entry()
        self.ltv_entry.insert(0, "70") 
        self.ltv_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="שכירות חודשית צפויה (₪):").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.rent_entry = new_entry()
        self.rent_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        self.skip_tax_var = tk.BooleanVar()
        self.tax_checkbox = ttk.Checkbutton(self.input_frame, text="בטל מס רכישה", variable=self.skip_tax_var)
        self.tax_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        r += 1

        self.include_tax_in_mortgage_var = tk.BooleanVar()
        self.include_tax_in_mortgage_checkbox = ttk.Checkbutton(self.input_frame, 
                                                                 text="כלול מס רכישה במשכנתא", 
                                                                 variable=self.include_tax_in_mortgage_var)
        self.include_tax_in_mortgage_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        r += 1

        self.manual_lawyer_fee_var = tk.BooleanVar()
        self.manual_lawyer_fee_checkbox = ttk.Checkbutton(self.input_frame, text="הזן עלות עו\"ד ידנית", variable=self.manual_lawyer_fee_var, command=self._toggle_lawyer_fee_entry)
        self.manual_lawyer_fee_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        self.lawyer_fee_manual_entry = new_entry()
        self.lawyer_fee_manual_entry.grid(row=r, column=1, sticky="w", pady=pady)
        self.lawyer_fee_manual_entry.config(state='disabled') 
        r += 1

        self.skip_broker_var = tk.BooleanVar() 
        self.manual_broker_fee_var = tk.BooleanVar()
        self.manual_broker_fee_checkbox = ttk.Checkbutton(self.input_frame, text="הזן עלות מתווך ידנית", variable=self.manual_broker_fee_var, command=self._toggle_broker_fee_entry)
        self.manual_broker_fee_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        self.broker_fee_manual_entry = new_entry()
        self.broker_fee_manual_entry.grid(row=r, column=1, sticky="w", pady=pady)
        self.broker_fee_manual_entry.config(state='disabled') 
        self.broker_checkbox = ttk.Checkbutton(self.input_frame, text="בטל עלות מתווך", variable=self.skip_broker_var)
        self.broker_checkbox.grid(row=r+1, column=0, sticky="w", padx=padx, pady=pady)
        r += 2 
        
        self.calculate_affordability_var = tk.BooleanVar()
        self.calculate_affordability_checkbox = ttk.Checkbutton(self.input_frame, 
                                                                 text="חשב מחיר נכס לפי הון עצמי (₪):", 
                                                                 variable=self.calculate_affordability_var, 
                                                                 command=self._toggle_affordability_calculation)
        self.calculate_affordability_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        self.available_funds_entry = new_entry()
        self.available_funds_entry.grid(row=r, column=1, sticky="w", pady=pady)
        self.available_funds_entry.config(state='disabled')
        r += 1

        self.rate_entries = []
        self.years_entries = []
        for i in range(3):
            ttk.Label(self.input_frame, text=f"ריבית שנתית (תרחיש {i+1}) %:").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
            rate_entry = new_entry()
            rate_entry.grid(row=r, column=1, sticky="w", pady=pady)
            self.rate_entries.append(rate_entry)
            r += 1

            ttk.Label(self.input_frame, text=f"שנים להחזר (תרחיש {i+1}):").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
            years_entry = new_entry()
            years_entry.grid(row=r, column=1, sticky="w", pady=pady)
            self.years_entries.append(years_entry)
            r += 1

        self.calculate_tab_button = ttk.Button(self.content_frame, text="חשב נכס זה", command=self.calculate) 
        self.calculate_tab_button.pack(pady=10)

        self.results_frame = ttk.Frame(self.content_frame)
        self.results_frame.pack(fill="x", expand=True, pady=10) 
        
        r_res = 0 

        self.affordable_price_label = ttk.Label(self.results_frame, text="", font=("Arial", 11, "bold"))
        self.affordable_price_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=(0, 5))
        r_res += 1

        self.tax_label = ttk.Label(self.results_frame, text="")
        self.tax_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=(10,2))
        r_res += 1

        self.downpayment_label = ttk.Label(self.results_frame, text="")
        self.downpayment_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
        r_res += 1

        self.loan_amount_label = ttk.Label(self.results_frame, text="")
        self.loan_amount_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
        r_res += 1

        self.lawyer_fee_label = ttk.Label(self.results_frame, text="")
        self.lawyer_fee_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
        r_res += 1

        self.broker_fee_label = ttk.Label(self.results_frame, text="")
        self.broker_fee_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
        r_res += 1

        self.total_funds_label = ttk.Label(self.results_frame, text="", font=("Arial", 11, "bold"))
        self.total_funds_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=(5,10))
        r_res += 1

        self.price_per_meter_label = ttk.Label(self.results_frame, text="")
        self.price_per_meter_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
        r_res += 1

        self.rent_comparison_labels = []
        for i in range(3):
            lbl = ttk.Label(self.results_frame, text="")
            lbl.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
            self.rent_comparison_labels.append(lbl)
            r_res += 1

        columns = ("loan", "rate", "years", "monthly", "interest", "total")
        self.table = ttk.Treeview(self.results_frame, columns=columns, show="headings", height=6) 
        self.table.grid(row=r_res, column=0, columnspan=2, sticky='nsew', pady=10, padx=padx)
        for col, title in zip(columns, ["סכום הלוואה (₪)", "ריבית שנתית (%)", "שנים להחזר", "תשלום חודשי (₪)", "סה\"כ ריבית (₪)", "סה\"כ תשלום כולל (₪)"]):
            self.table.heading(col, text=title)
            self.table.column(col, width=150, anchor="center") 
        r_res += 1

        self.results_frame.grid_rowconfigure(r_res, weight=1)
        self.results_frame.grid_columnconfigure(1, weight=1) 

        self.figure_list = []
        self.ax_list = []
        self.canvas_list = []
        self.temp_image_paths = [] 

        for i in range(3):
            fig = plt.Figure(figsize=(5, 2.5), dpi=100) 
            ax = fig.add_subplot(111)
            canvas = FigureCanvasTkAgg(fig, self.results_frame) 
            canvas.get_tk_widget().grid(row=r_res+i, column=0, columnspan=2, pady=3, sticky='nsew')
            self.figure_list.append(fig)
            self.ax_list.append(ax)
            self.canvas_list.append(canvas)

        self.df_list = [None, None, None] 

        self.calculated_results = {}
        self.loan_scenarios_data = [] 
        self.loan_scenarios_rent_comparison = []

        self.content_frame.bind('<Configure>', self._on_frame_configure)

        self.export_pdf_button = ttk.Button(self.content_frame, text="ייצוא ל-PDF", command=self.export_to_pdf)
        self.export_pdf_button.pack(pady=10)

    def _on_frame_configure(self, event=None):
        """Update the scrollregion of the canvas based on the content frame size."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        """Adjust content frame width to match canvas width when canvas resizes."""
        canvas_width = event.width
        # This will adjust the width of the window *inside* the canvas
        self.canvas.itemconfig(self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw"), width=canvas_width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_mousewheel_up(self, event):
        self.canvas.yview_scroll(-1, "units")

    def _on_mousewheel_down(self, event):
        self.canvas.yview_scroll(1, "units")

    def _toggle_lawyer_fee_entry(self):
        if self.manual_lawyer_fee_var.get():
            self.lawyer_fee_manual_entry.config(state='normal')
        else:
            self.lawyer_fee_manual_entry.config(state='disabled')
            self.lawyer_fee_manual_entry.delete(0, tk.END) 

    def _toggle_broker_fee_entry(self):
        if self.manual_broker_fee_var.get():
            self.broker_fee_manual_entry.config(state='normal')
            self.broker_checkbox.config(state='disabled') 
            self.skip_broker_var.set(False) 
        else:
            self.broker_fee_manual_entry.config(state='disabled')
            self.broker_fee_manual_entry.delete(0, tk.END) 
            self.broker_checkbox.config(state='normal') 

    def _toggle_affordability_calculation(self):
        if self.calculate_affordability_var.get():
            self.available_funds_entry.config(state='normal')
            self.price_entry.config(state='disabled')
        else:
            self.available_funds_entry.config(state='disabled')
            self.available_funds_entry.delete(0, tk.END)
            self.price_entry.config(state='normal')

    def clear_results(self):
        self.affordable_price_label.config(text="")
        self.tax_label.config(text="")
        self.downpayment_label.config(text="")
        self.loan_amount_label.config(text="")
        self.lawyer_fee_label.config(text="")
        self.broker_fee_label.config(text="")
        self.total_funds_label.config(text="")
        self.price_per_meter_label.config(text="")
        for lbl in self.rent_comparison_labels:
            lbl.config(text="")
        self.table.delete(*self.table.get_children())
        for ax in self.ax_list:
            ax.clear()
        for canvas in self.canvas_list:
            canvas.draw()
        self.df_list = [None, None, None] 
        self.calculated_results = {}
        self.loan_scenarios_data = [] 
        self.loan_scenarios_rent_comparison = []


    def calculate(self):
        self.clear_results() 

        is_active_tab = (self.idx == self.frame.master.index(self.frame)) if hasattr(self.frame.master, 'index') else False

        try:
            price = 0.0
            loan_amount = 0.0
            down_payment = 0.0
            purchase_tax = 0.0
            lawyer_fee = 0.0
            broker_fee = 0.0
            total_needed = 0.0

            ltv_str = self.ltv_entry.get()
            if not ltv_str:
                if is_active_tab:
                    show_error_with_copy("קלט חסר", "יש להזין אחוז מימון (LTV).", parent=self.root)
                return False
            ltv = float(ltv_str)
            if not (0 <= ltv <= 100):
                if is_active_tab:
                    show_error_with_copy("קלט לא חוקי", "אחוז מימון (LTV) חייב להיות בין 0 ל-100.", parent=self.root)
                return False

            area_str = self.area_entry.get()
            area = float(area_str) if area_str else None
            if area is not None and area <= 0:
                if is_active_tab:
                    show_error_with_copy("קלט לא חוקי", "שטח המטר המרובע חייב להיות מספר חיובי.", parent=self.root)
                return False

            rent_str = self.rent_entry.get()
            rent = float(rent_str) if rent_str else None
            if rent is not None and rent < 0:
                if is_active_tab:
                    show_error_with_copy("קלט לא חוקי", "שכירות חודשית צפויה אינה יכולה להיות שלילית.", parent=self.root)
                return False

            if self.calculate_affordability_var.get():
                available_funds_str = self.available_funds_entry.get()
                if not available_funds_str:
                    if is_active_tab:
                        show_error_with_copy("קלט חסר", "יש להזין את סכום הכסף הפנוי.", parent=self.root)
                    return False
                available_funds = float(available_funds_str)
                if available_funds <= 0:
                    if is_active_tab:
                        show_error_with_copy("קלט לא חוקי", "סכום הכסף הפנוי חייב להיות חיובי.", parent=self.root)
                    return False

                estimated_price = available_funds / ((1 - ltv / 100) + LAWYER_FEE_RATE + BROKER_FEE_RATE)
                
                tolerance = 1.0
                max_iterations = 100
                current_iteration = 0
                
                

                while current_iteration < max_iterations:
                    current_down_payment_ratio = (1 - ltv / 100)
                    if self.include_tax_in_mortgage_var.get():
                        include_tax_ind = True
                        current_down_payment_from_price = (estimated_price ) * current_down_payment_ratio
                        current_purchase_tax = 0  if self.skip_tax_var.get() else calculate_purchase_tax(estimated_price)
                        
                    if self.manual_lawyer_fee_var.get():
                        lawyer_fee_val = float(self.lawyer_fee_manual_entry.get()) if self.lawyer_fee_manual_entry.get() else 0
                        current_lawyer_fee = lawyer_fee_val
                    else:
                        current_lawyer_fee = estimated_price * LAWYER_FEE_RATE

                    if self.skip_broker_var.get():
                        current_broker_fee = 0
                    elif self.manual_broker_fee_var.get():
                        broker_fee_val = float(self.broker_fee_manual_entry.get()) if self.broker_fee_manual_entry.get() else 0
                        current_broker_fee = broker_fee_val
                    else:
                        current_broker_fee = estimated_price * BROKER_FEE_RATE

                    if self.include_tax_in_mortgage_var.get():


                # loan_amount_f=(price + purchase_tax)
                # loan_amount = loan_amount_f* (ltv / 100)
                # down_payment =( price + purchase_tax )* ((100-ltv) / 100)
                        current_price_tax_f=calculate_purchase_tax(estimated_price)
                        current_loan_amount_for_affordability_f=int(estimated_price + current_price_tax_f)
                        current_loan_amount_for_affordability = (current_loan_amount_for_affordability_f * (ltv / 100)) 
                        # current_down_payment_for_affordability = (estimated_price + current_purchase_tax) - current_loan_amount_for_affordability
                        current_down_payment_for_affordability = (estimated_price + calculate_purchase_tax(estimated_price))* ((100-ltv) / 100)
                        current_total_funds_needed = current_down_payment_for_affordability + current_lawyer_fee + current_broker_fee
                    else:
                        current_total_funds_needed = current_down_payment_from_price + current_purchase_tax + current_lawyer_fee + current_broker_fee

                    diff = available_funds - current_total_funds_needed

                    if abs(diff) < tolerance:
                        price = estimated_price
                        break
                    
                    estimated_price += diff * 0.5 

                    current_iteration += 1
                
                if current_iteration == max_iterations:
                    if is_active_tab:
                        show_error_with_copy("אזהרת חישוב", "לא ניתן למצוא מחיר נכס מדויק עבור ההון העצמי הנתון לאחר מספר רב של ניסיונות. ייתכן שהסכום המחושב הוא קירוב.", parent=self.root)
                    price = estimated_price 

                self.price_entry.config(state='disabled')
                self.price_entry.delete(0, tk.END)
                self.price_entry.insert(0, f"{price:,.0f}")
                self.affordable_price_label.config(text=f"מחיר הנכס המקסימלי שניתן לרכוש: {price:,.0f} ₪")

            else: 
                price_str = self.price_entry.get()
                if not price_str:
                    if is_active_tab:
                        show_error_with_copy("קלט חסר", "יש להזין מחיר דירה.", parent=self.root)
                    return False 
                price = float(price_str)
                if price <= 0:
                    if is_active_tab:
                        show_error_with_copy("קלט לא חוקי", "מחיר הדירה חייב להיות מספר חיובי.", parent=self.root)
                    return False
                self.price_entry.config(state='normal')
                self.affordable_price_label.config(text="") 

            purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
            
            if self.manual_lawyer_fee_var.get():
                try:
                    lawyer_fee = float(self.lawyer_fee_manual_entry.get())
                    if lawyer_fee < 0:
                        if is_active_tab:
                            show_error_with_copy("קלט לא חוקי", "עלות עו\"ד ידנית אינה יכולה להיות שלילית.", parent=self.root)
                        return False
                except ValueError:
                    if is_active_tab:
                        show_error_with_copy("שגיאת קלט", "עלות עו\"ד ידנית חייבת להיות מספר.", parent=self.root)
                    return False
            else:
                lawyer_fee = estimate_lawyer_fee(price)

            if self.skip_broker_var.get():
                broker_fee = 0
            elif self.manual_broker_fee_var.get():
                try:
                    broker_fee = float(self.broker_fee_manual_entry.get())
                    if broker_fee < 0:
                        if is_active_tab:
                            show_error_with_copy("קלט לא חוקי", "עלות מתווך ידנית אינה יכולה להיות שלילית.", parent=self.root)
                        return False
                except ValueError:
                    if is_active_tab:
                        show_error_with_copy("שגיאת קלט", "עלות מתווך ידנית חייבת להיות מספר.", parent=self.root)
                    return False
            else:
                broker_fee = estimate_broker_fee(price)

            base_loan_amount = price * (ltv / 100)
            if self.include_tax_in_mortgage_var.get():
                loan_amount_f=(price + purchase_tax)
                loan_amount = loan_amount_f* (ltv / 100)
                down_payment =( price + purchase_tax )* ((100-ltv) / 100)
            else:
                loan_amount = base_loan_amount
                down_payment = (price - base_loan_amount) + purchase_tax

            total_needed = down_payment + lawyer_fee + broker_fee

            if self.calculate_affordability_var.get():
                total_needed = available_funds 

            self.calculated_results = {
                "purchase_tax": purchase_tax,
                "down_payment": down_payment,
                "loan_amount": loan_amount, 
                "lawyer_fee": lawyer_fee,
                "broker_fee": broker_fee,
                "total_needed": total_needed,
                "price_per_meter": price / area if area is not None and area > 0 else None,
                "rent": rent,
                "input_price": self.price_entry.get(), 
                "calculated_price": price, 
                "input_area": area_str,
                "input_ltv": ltv_str,
                "input_rent": rent_str,
                "input_skip_tax": self.skip_tax_var.get(),
                "input_include_tax_in_mortgage": self.include_tax_in_mortgage_var.get(), 
                "input_skip_broker": self.skip_broker_var.get(),
                "input_manual_lawyer_fee": self.manual_lawyer_fee_var.get(), 
                "input_lawyer_fee_manual_value": self.lawyer_fee_manual_entry.get(),
                "input_manual_broker_fee": self.manual_broker_fee_var.get(),
                "input_broker_fee_manual_value": self.broker_fee_manual_entry.get(),
                "input_calculate_affordability": self.calculate_affordability_var.get(),
                "input_available_funds": self.available_funds_entry.get(),
                "input_rates": [], 
                "input_years": [], 
                "input_alias": self.alias_entry.get(),
                "input_link": self.link_entry.get(),
            }
            self.loan_scenarios_data = [] 
            self.loan_scenarios_rent_comparison = []

            rates = []
            years = []
            valid_scenarios_count = 0
            for i in range(3):
                rate_val = self.rate_entries[i].get()
                years_val = self.years_entries[i].get()
                
                current_rate = None
                current_years = None

                if rate_val and years_val: 
                    try:
                        current_rate = float(rate_val)
                        if current_rate < 0:
                            if is_active_tab:
                                show_error_with_copy("קלט לא חוקי", f"ריבית שנתית (תרחיש {i+1}) אינה יכולה להיות שלילית.", parent=self.root)
                            return False
                    except ValueError:
                        if is_active_tab:
                            show_error_with_copy("שגיאת קלט", f"ריבית שנתית (תרחיש {i+1}) חייבת להיות מספר.", parent=self.root)
                        return False
                    
                    try:
                        current_years = int(years_val)
                        if current_years <= 0:
                            if is_active_tab:
                                show_error_with_copy("קלט לא חוקי", f"שנים להחזר (תרחיש {i+1}) חייבות להיות מספר חיובי שלם.", parent=self.root)
                            return False
                    except ValueError:
                        if is_active_tab:
                            show_error_with_copy("שגיאת קלט", f"שנים להחזר (תרחיש {i+1}) חייבות להיות מספר שלם.", parent=self.root)
                        return False
                    valid_scenarios_count += 1
                
                rates.append(current_rate)
                years.append(current_years)

            if valid_scenarios_count == 0:
                if is_active_tab:
                    show_error_with_copy("אין נתונים לחישוב", "אנא הזן/י לפחות ריבית שנתית אחת ושנים להחזר עבור תרחיש.", parent=self.root)
                return False 
            
            self.calculated_results["input_rates"] = rates
            self.calculated_results["input_years"] = years

            self.tax_label.config(text=f"מס רכישה משוער: {purchase_tax:,.0f} ₪")
            self.downpayment_label.config(text=f"הון עצמי נדרש: {down_payment:,.0f} ₪")
            self.loan_amount_label.config(text=f"סכום הלוואה מהבנק: {loan_amount:,.0f} ₪")
            self.lawyer_fee_label.config(text=f"עלות עורך דין משוערת: {lawyer_fee:,.0f} ₪")
            self.broker_fee_label.config(text=f"עלות מתווך משוערת: {broker_fee:,.0f} ₪")
            self.total_funds_label.config(text=f"סה\"כ הון דרוש: {total_needed:,.0f} ₪")

            if area is not None and area > 0: 
                self.price_per_meter_label.config(text=f"מחיר למטר מרובע: {price / area:,.2f} ₪")
            else:
                self.price_per_meter_label.config(text="") 

            self.table.delete(*self.table.get_children()) 

            for p in self.temp_image_paths:
                try:
                    import os
                    os.remove(p)
                except OSError:
                    pass
            self.temp_image_paths = []

            for i in range(3):
                if rates[i] is not None and years[i] is not None:
                    df = generate_amortization_df(loan_amount, rates[i], years[i])
                    self.df_list[i] = df
                    
                    if not df.empty:
                        total_interest = df["ריבית"].sum()
                        total_payment_sum_from_df = df["תשלום חודשי"].sum()
                        
                        initial_monthly_payment_for_scenario = calculate_monthly_payment(loan_amount, rates[i], years[i])

                        table_row_data = (
                            f"{loan_amount:,.0f}",
                            f"{rates[i]:.2f}",
                            f"{years[i]}",
                            f"{initial_monthly_payment_for_scenario:,.0f}", 
                            f"{total_interest:,.0f}",
                            f"{total_payment_sum_from_df:,.0f}", 
                        )
                        self.table.insert("", "end", values=table_row_data)
                        self.loan_scenarios_data.append({
                            "תרחיש": f"תרחיש {i+1}",
                            "סכום הלוואה (₪)": f"{loan_amount:,.0f}",
                            "ריבית שנתית (%)": f"{rates[i]:.2f}",
                            "שנים להחזר": f"{years[i]}",
                            "תשלום חודשי (₪)": f"{initial_monthly_payment_for_scenario:,.0f}",
                            "סה\"כ ריבית (₪)": f"{total_interest:,.0f}",
                            "סה\"כ תשלום כולל (₪)": f"{total_payment_sum_from_df:,.0f}"
                        })

                        rent_compare_str = ""
                        if rent is not None: 
                            ratio = rent / initial_monthly_payment_for_scenario if initial_monthly_payment_for_scenario != 0 else 0
                            rent_compare_str = f"שכירות צפויה: {rent:,.0f} ₪ | תשלום חודשי ראשוני: {initial_monthly_payment_for_scenario:,.0f} ₪ | יחס שכירות/תשלום: {ratio:.2f}"
                            self.rent_comparison_labels[i].config(text=rent_compare_str)
                        else:
                            self.rent_comparison_labels[i].config(text="") 
                        self.loan_scenarios_rent_comparison.append(rent_compare_str)
                        
                        ax = self.ax_list[i]
                        ax.clear()
                        ax.plot(df["חודש"], df["קרן"], label="קרן", color="green")
                        ax.plot(df["חודש"], df["ריבית"], label="ריבית", color="red")
                        ax.set_title(f"תרחיש {i+1} - פירוט תשלומים חודשיים", fontsize=9)
                        ax.set_xlabel("חודש", fontsize=8)
                        ax.set_ylabel("₪", fontsize=8)
                        ax.legend(fontsize=7)
                        ax.grid(True)
                        ax.set_xlim(left=1)
                        ax.tick_params(axis='both', which='major', labelsize=7) 
                        self.figure_list[i].tight_layout() 
                        self.canvas_list[i].draw()
                    else:
                        self.df_list[i] = None
                        self.table.insert("", "end", values=("אין נתונים עבור תרחיש זה",) * 6)
                        self.rent_comparison_labels[i].config(text="")
                        self.ax_list[i].clear()
                        self.canvas_list[i].draw()
                        self.loan_scenarios_data.append({}) 
                        self.loan_scenarios_rent_comparison.append("אין נתוני השוואת שכירות עבור תרחיש זה")
                else:
                    self.df_list[i] = None
                    self.table.insert("", "end", values=("אין נתונים עבור תרחיש זה (חסר ריבית/שנים)",) * 6)
                    self.rent_comparison_labels[i].config(text="")
                    self.ax_list[i].clear()
                    self.canvas_list[i].draw()
                    self.loan_scenarios_data.append({}) 
                    self.loan_scenarios_rent_comparison.append("אין נתוני השוואת שכירות עבור תרחיש זה")
            
            self._on_frame_configure()

            return True 

        except ValueError as e:
            if is_active_tab:
                show_error_with_copy("שגיאת קלט", f"שגיאה בנתונים: {e}\nאנא ודא/י שכל השדות המספריים מולאו נכונה.", parent=self.root)
            return False
        except Exception as e:
            if is_active_tab:
                show_error_with_copy("שגיאה כללית", f"אירעה שגיאה בלתי צפויה: {e}", parent=self.root)
            return False

    def export_to_pdf(self):
        if not self.calculate():
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                filetypes=[("PDF files", "*.pdf")],
                                                title="שמור דוח נכס (PDF)")
        if not filepath:
            return

        try:
            doc = SimpleDocTemplate(filepath, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36) # Added margins
            
            # Use the Hebrew styles defined globally
            styles = getSampleStyleSheet()
            # Ensure these styles are added if they were not created due to font loading errors
            # (they are already added in the global scope if font loading succeeded)
            if 'Hebrew' not in styles:
                styles.add(heb_style)
            if 'HebrewHeading' not in styles:
                styles.add(heb_heading_style)
            if 'HebrewSubHeading' not in styles:
                styles.add(heb_subheading_style)
            
            story = []

            # --- Title ---
            alias = self.calculated_results.get("input_alias", f"נכס {self.idx + 1}")
            story.append(Paragraph(f"<b>דוח נכס: {alias}</b>", styles['HebrewHeading']))
            story.append(Spacer(1, 0.2 * inch))

            # --- Input Data Section ---
            story.append(Paragraph("<b>פרטי קלט:</b>", styles['HebrewSubHeading']))
            story.append(Spacer(1, 0.1 * inch))

            input_data = [
                ("קישור:", self.calculated_results.get("input_link", "")),
                ("מחיר דירה (₪):", f"{self.calculated_results.get('calculated_price', 0):,.0f}"),
                ("מטר מרובע (שטח):", self.calculated_results.get("input_area", "")),
                ("אחוז מימון (LTV) %:", self.calculated_results.get("input_ltv", "")),
                ("שכירות חודשית צפויה (₪):", self.calculated_results.get("input_rent", "")),
                ("בטל מס רכישה:", "כן" if self.calculated_results.get("input_skip_tax") else "לא"),
                ("כלול מס רכישה במשכנתא:", "כן" if self.calculated_results.get("input_include_tax_in_mortgage") else "לא"),
            ]
            
            if self.calculated_results.get("input_manual_lawyer_fee"):
                input_data.append(("הזן עלות עו\"ד ידנית:", self.calculated_results.get("input_lawyer_fee_manual_value", "")))
            else:
                input_data.append(("עלות עו\"ד משוערת (% מהמחיר):", f"{LAWYER_FEE_RATE*100:.0f}%"))

            if self.calculated_results.get("input_manual_broker_fee"):
                 input_data.append(("הזן עלות מתווך ידנית:", self.calculated_results.get("input_broker_fee_manual_value", "")))
            elif self.calculated_results.get("input_skip_broker"):
                input_data.append(("בטל עלות מתווך:", "כן"))
            else:
                input_data.append(("עלות מתווך משוערת (% מהמחיר):", f"{BROKER_FEE_RATE*100:.0f}%"))
            
            if self.calculated_results.get("input_calculate_affordability"):
                input_data.append(("חשב מחיר נכס לפי הון עצמי (₪):", self.calculated_results.get("input_available_funds", "")))

            # Use the Hebrew style for Paragraphs in the table
            input_table_data = [[Paragraph(f"<b>{k}</b>", styles['Hebrew']), Paragraph(str(v), styles['Hebrew'])] for k, v in input_data]
            
            # Adjust colWidths to prevent cutting and fit content
            table_col_widths = [doc.width * 0.4, doc.width * 0.6] # Allocate width dynamically
            input_table = Table(input_table_data, colWidths=table_col_widths)
            input_table.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'), # Align right for Hebrew
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('FONTNAME', (0,0), (0,-1), 'DejaVuSans-Bold'), # This now refers to the registered bold font
                ('FONTNAME', (1,0), (1,-1), 'DejaVuSans'), # This now refers to the registered normal font
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                ('GRID', (0,0), (-1,-1), 0.25, colors.black),
                ('BACKGROUNDS', (0,0), (-1,-1), [colors.HexColor('#F0F8FF'), None]), # Light blue for alternating rows
            ]))
            story.append(input_table)
            story.append(Spacer(1, 0.3 * inch))

            # --- Calculation Summary ---
            story.append(Paragraph("<b>סיכום חישובים:</b>", styles['HebrewSubHeading']))
            story.append(Spacer(1, 0.1 * inch))

            summary_data = [
                ("מס רכישה משוער:", f"{self.calculated_results.get('purchase_tax', 0):,.0f} ₪"),
                ("הון עצמי נדרש:", f"{self.calculated_results.get('down_payment', 0):,.0f} ₪"),
                ("סכום הלוואה מהבנק:", f"{self.calculated_results.get('loan_amount', 0):,.0f} ₪"),
                ("עלות עורך דין משוערת:", f"{self.calculated_results.get('lawyer_fee', 0):,.0f} ₪"),
                ("עלות מתווך משוערת:", f"{self.calculated_results.get('broker_fee', 0):,.0f} ₪"),
                ("סה\"כ הון דרוש:", f"{self.calculated_results.get('total_needed', 0):,.0f} ₪"),
            ]
            if self.calculated_results.get("price_per_meter") is not None:
                summary_data.append(("מחיר למטר מרובע:", f"{self.calculated_results.get('price_per_meter', 0):,.2f} ₪"))

            summary_table_data = [[Paragraph(f"<b>{k}</b>", styles['Hebrew']), Paragraph(str(v), styles['Hebrew'])] for k, v in summary_data]
            summary_table = Table(summary_table_data, colWidths=table_col_widths)
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'), # Align right for Hebrew
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('FONTNAME', (0,0), (0,-1), 'DejaVuSans-Bold'), # This too
                ('FONTNAME', (1,0), (1,-1), 'DejaVuSans'), # And this
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                ('GRID', (0,0), (-1,-1), 0.25, colors.black),
                ('BACKGROUNDS', (0,0), (-1,-1), [colors.HexColor('#F0F8FF'), None]),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 0.3 * inch))

            # --- Loan Scenarios Table ---
            story.append(Paragraph("<b>תרחישי הלוואה:</b>", styles['HebrewSubHeading']))
            story.append(Spacer(1, 0.1 * inch))

            loan_table_headers = ["תרחיש", "סכום הלוואה (₪)", "ריבית שנתית (%)", "שנים להחזר", "תשלום חודשי (₪)", "סה\"כ ריבית (₪)", "סה\"כ תשלום כולל (₪)"]
            
            # Wrap headers in Paragraphs for font styling
            loan_table_data = [[Paragraph(header, styles['Hebrew']) for header in loan_table_headers]]
            
            for i, scenario in enumerate(self.loan_scenarios_data):
                if scenario:
                    # Wrap each cell's content in Paragraph for font styling
                    loan_table_data.append([
                        Paragraph(scenario.get("תרחיש", ""), styles['Hebrew']),
                        Paragraph(scenario.get("סכום הלוואה (₪)", ""), styles['Hebrew']),
                        Paragraph(scenario.get("ריבית שנתית (%)", ""), styles['Hebrew']),
                        Paragraph(scenario.get("שנים להחזר", ""), styles['Hebrew']),
                        Paragraph(scenario.get("תשלום חודשי (₪)", ""), styles['Hebrew']),
                        Paragraph(scenario.get("סה\"כ ריבית (₪)", ""), styles['Hebrew']),
                        Paragraph(scenario.get("סה\"כ תשלום כולל (₪)", ""), styles['Hebrew'])
                    ])
            
            if len(loan_table_data) > 1:
                # Calculate optimal column widths based on content or fixed proportions
                # Adjust colWidths to fit content. A4 width is ~595 points, effective width ~523 points.
                # 523 / 7 columns ~= 75 points per column. Let's make it a bit more flexible.
                col_widths = [doc.width * 0.12, doc.width * 0.17, doc.width * 0.12, doc.width * 0.1, doc.width * 0.17, doc.width * 0.16, doc.width * 0.16] # Adjusted widths
                
                loan_table = Table(loan_table_data, colWidths=col_widths)
                loan_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ADD8E6')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('FONTNAME', (0,0), (-1,0), 'DejaVuSans-Bold'), # Ensure header font is bold
                    ('FONTSIZE', (0,0), (-1,0), 9), # Smaller font for table headers
                    ('BOTTOMPADDING', (0,0), (-1,0), 6),
                    ('BACKGROUNDS', (0,1), (-1,-1), [colors.beige, colors.white]),
                    ('GRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('FONTNAME', (0,1), (-1,-1), 'DejaVuSans'), # Regular font for table data
                    ('FONTSIZE', (0,1), (-1,-1), 8), # Smaller font for table data
                ]))
                story.append(loan_table)
                story.append(Spacer(1, 0.3 * inch))
            else:
                story.append(Paragraph("אין נתוני הלוואה לתרחישים.", styles['Hebrew']))
                story.append(Spacer(1, 0.3 * inch))

            # --- Rent Comparison ---
            if any(self.loan_scenarios_rent_comparison):
                story.append(Paragraph("<b>השוואת שכירות:</b>", styles['HebrewSubHeading']))
                story.append(Spacer(1, 0.1 * inch))
                for rent_comp_str in self.loan_scenarios_rent_comparison:
                    if rent_comp_str:
                        story.append(Paragraph(rent_comp_str, styles['Hebrew']))
                        story.append(Spacer(1, 0.05 * inch))
                story.append(Spacer(1, 0.3 * inch))

            # --- Amortization Graphs ---
            story.append(Paragraph("<b>גרפי פירעון:</b>", styles['HebrewSubHeading']))
            story.append(Spacer(1, 0.1 * inch))

            for i, fig in enumerate(self.figure_list):
                # Check if df_list[i] is not None AND not empty
                if self.df_list[i] is not None and not self.df_list[i].empty: 
                    buf = io.BytesIO()
                    fig.savefig(buf, format='png', dpi=200, bbox_inches='tight') # Increased DPI for better quality
                    buf.seek(0)
                    
                    img = RLImage(buf)
                    
                    # Calculate aspect ratio to fit within page width
                    img_width, img_height = img.drawWidth, img.drawHeight
                    aspect_ratio = img_height / img_width
                    
                    # Target width for the image (e.g., 7 inches, leaving margins)
                    desired_width = 7 * inch # Adjusted to fit page width with margins
                    desired_height = desired_width * aspect_ratio

                    # If the image is too tall, scale down based on height as well
                    if desired_height > (A4[1] - (36*2 + 1*inch)): # A4 height - top/bottom margins - some space for title/text
                         desired_height = (A4[1] - (36*2 + 1*inch))
                         desired_width = desired_height / aspect_ratio
                    
                    img.drawWidth = desired_width
                    img.drawHeight = desired_height
                    
                    # Add PageBreak if graph won't fit on current page
                    current_y_pos = doc.height - doc.topMargin - sum(el.wrapOn(doc, doc.width, doc.height)[1] for el in story[-2:]) # Estimate height of last elements
                    if (current_y_pos - desired_height - (0.5*inch)) < doc.bottomMargin and i > 0: # Check if there's enough space
                        story.append(PageBreak())
                        story.append(Paragraph(f"<b>גרפי פירעון (המשך):</b>", styles['HebrewSubHeading']))
                        story.append(Spacer(1, 0.1 * inch))

                    story.append(Paragraph(f"<b>תרחיש {i+1}</b>", styles['Hebrew']))
                    story.append(img)
                    story.append(Spacer(1, 0.2 * inch))
            
            doc.build(story)
            show_error_with_copy("ייצוא ל-PDF", "הדוח נשמר בהצלחה כקובץ PDF.", parent=self.root)

        except Exception as e:
            show_error_with_copy("שגיאת ייצוא ל-PDF", f"אירעה שגיאה בעת ייצוא ל-PDF: {e}", parent=self.root)


class MortgageCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("מחשבון נדל\"ן מקיף")
        self.root.geometry("1200x900") 

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.property_tabs = []
        self.add_tab()

        menu_bar = tk.Menu(root)
        root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="קובץ", menu=file_menu)
        file_menu.add_command(label="הוסף נכס חדש", command=self.add_tab)
        file_menu.add_command(label="שמור נתונים (Excel)", command=self.save_data)
        file_menu.add_command(label="טען נתונים (Excel)", command=self.load_data)
        file_menu.add_separator()
        file_menu.add_command(label="יציאה", command=root.quit)

    def add_tab(self):
        idx = len(self.property_tabs)
        new_tab = PropertyTab(self.notebook, idx, self.root) 
        self.property_tabs.append(new_tab)
        self.notebook.add(new_tab.frame, text=f"נכס {idx + 1}")
        self.notebook.select(new_tab.frame) 

    def save_data(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="שמור נתוני נכסים")
        if not filepath:
            return

        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                summary_data = []
                for idx, prop_tab in enumerate(self.property_tabs):
                    prop_tab.calculate() 
                    
                    results = prop_tab.calculated_results
                    loan_scenarios = prop_tab.loan_scenarios_data
                    rent_comparisons = prop_tab.loan_scenarios_rent_comparison

                    alias = results.get("input_alias", f"נכס {idx + 1}")
                    link = results.get("input_link", "")

                    summary_row = {
                        "Alias": alias,
                        "Link": link,
                        "מחיר דירה (₪)": results.get("calculated_price"),
                        "מטר מרובע (שטח)": results.get("input_area"),
                        "אחוז מימון (LTV) %": results.get("input_ltv"),
                        "שכירות חודשית צפויה (₪)": results.get("input_rent"),
                        "בטל מס רכישה": "כן" if results.get("input_skip_tax") else "לא",
                        "כלול מס רכישה במשכנתא": "כן" if results.get("input_include_tax_in_mortgage") else "לא",
                        "הזן עלות עו\"ד ידנית": "כן" if results.get("input_manual_lawyer_fee") else "לא",
                        "עלות עו\"ד ידנית": results.get("input_lawyer_fee_manual_value"),
                        "הזן עלות מתווך ידנית": "כן" if results.get("input_manual_broker_fee") else "לא",
                        "עלות מתווך ידנית": results.get("input_broker_fee_manual_value"),
                        "בטל עלות מתווך": "כן" if results.get("input_skip_broker") else "לא",
                        "חשב מחיר נכס לפי הון עצמי": "כן" if results.get("input_calculate_affordability") else "לא",
                        "הון עצמי זמין (₪)": results.get("input_available_funds"),
                        "מס רכישה משוער (₪)": results.get("purchase_tax"),
                        "הון עצמי נדרש (₪)": results.get("down_payment"),
                        "סכום הלוואה מהבנק (₪)": results.get("loan_amount"),
                        "עלות עורך דין משוערת (₪)": results.get("lawyer_fee"),
                        "עלות מתווך משוערת (₪)": results.get("broker_fee"),
                        "סה\"כ הון דרוש (₪)": results.get("total_needed"),
                        "מחיר למטר מרובע (₪)": results.get("price_per_meter"),
                    }
                    
                    for i, scenario in enumerate(loan_scenarios):
                        prefix = f"תרחיש {i+1} - "
                        summary_row[prefix + "סכום הלוואה (₪)"] = scenario.get("סכום הלוואה (₪)")
                        summary_row[prefix + "ריבית שנתית (%)"] = scenario.get("ריבית שנתית (%)")
                        summary_row[prefix + "שנים להחזר"] = scenario.get("שנים להחזר")
                        summary_row[prefix + "תשלום חודשי (₪)"] = scenario.get("תשלום חודשי (₪)")
                        summary_row[prefix + "סה\"כ ריבית (₪)"] = scenario.get("סה\"כ ריבית (₪)")
                        summary_row[prefix + "סה\"כ תשלום כולל (₪)"] = scenario.get("סה\"כ תשלום כולל (₪)")
                        
                    for i, rent_comp in enumerate(rent_comparisons):
                        summary_row[f"תרחיש {i+1} - השוואת שכירות"] = rent_comp

                    summary_data.append(summary_row)

                    for i, df in enumerate(prop_tab.df_list):
                        if df is not None and not df.empty:
                            sheet_name = f"{alias}_תרחיש_{i+1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)

                pd.DataFrame(summary_data).to_excel(writer, sheet_name="סיכום נכסים", index=False)
            
            show_error_with_copy("שמירה בוצעה", "הנתונים נשמרו בהצלחה לקובץ Excel.", parent=self.root)

        except Exception as e:
            show_error_with_copy("שגיאה בשמירה", f"אירעה שגיאה בעת שמירת הנתונים: {e}", parent=self.root)

    def load_data(self):
        filepath = filedialog.askopenfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="טען נתוני נכסים")
        if not filepath:
            return

        try:
            xls = pd.ExcelFile(filepath)
            
            for _ in range(len(self.property_tabs)):
                self.notebook.forget(0)
            self.property_tabs = []

            if "סיכום נכסים" not in xls.sheet_names:
                show_error_with_copy("שגיאה בטעינה", "קובץ Excel אינו מכיל גיליון 'סיכום נכסים'.", parent=self.root)
                return

            summary_df = pd.read_excel(xls, sheet_name="סיכום נכסים")

            for index, row in summary_df.iterrows():
                self.add_tab()
                current_tab = self.property_tabs[-1]

                current_tab.alias_entry.delete(0, tk.END)
                current_tab.alias_entry.insert(0, row.get("Alias", ""))
                
                current_tab.link_entry.delete(0, tk.END)
                current_tab.link_entry.insert(0, row.get("Link", ""))

                current_tab.price_entry.delete(0, tk.END)
                if not row.get("חשב מחיר נכס לפי הון עצמי"): 
                    price_val = row.get("מחיר דירה (₪)")
                    if pd.notna(price_val):
                        current_tab.price_entry.insert(0, str(int(price_val)))

                current_tab.area_entry.delete(0, tk.END)
                area_val = row.get("מטר מרובע (שטח)")
                if pd.notna(area_val):
                    current_tab.area_entry.insert(0, str(int(area_val)))

                current_tab.ltv_entry.delete(0, tk.END)
                ltv_val = row.get("אחוז מימון (LTV) %")
                if pd.notna(ltv_val):
                    current_tab.ltv_entry.insert(0, str(int(ltv_val)))

                current_tab.rent_entry.delete(0, tk.END)
                rent_val = row.get("שכירות חודשית צפויה (₪)")
                if pd.notna(rent_val):
                    current_tab.rent_entry.insert(0, str(int(rent_val)))

                current_tab.skip_tax_var.set(row.get("בטל מס רכישה") == "כן")
                current_tab.include_tax_in_mortgage_var.set(row.get("כלול מס רכישה במשכנתא") == "כן")

                manual_lawyer = (row.get("הזן עלות עו\"ד ידנית") == "כן")
                current_tab.manual_lawyer_fee_var.set(manual_lawyer)
                current_tab._toggle_lawyer_fee_entry() 
                if manual_lawyer and pd.notna(row.get("עלות עו\"ד ידנית")):
                    current_tab.lawyer_fee_manual_entry.delete(0, tk.END)
                    current_tab.lawyer_fee_manual_entry.insert(0, str(int(row["עלות עו\"ד ידנית"])))

                manual_broker = (row.get("הזן עלות מתווך ידנית") == "כן")
                current_tab.manual_broker_fee_var.set(manual_broker)
                current_tab._toggle_broker_fee_entry() 
                if manual_broker and pd.notna(row.get("עלות מתווך ידנית")):
                    current_tab.broker_fee_manual_entry.delete(0, tk.END)
                    current_tab.broker_fee_manual_entry.insert(0, str(int(row["עלות מתווך ידנית"])))
                current_tab.skip_broker_var.set(row.get("בטל עלות מתווך") == "כן")

                calc_afford = (row.get("חשב מחיר נכס לפי הון עצמי") == "כן")
                current_tab.calculate_affordability_var.set(calc_afford)
                current_tab._toggle_affordability_calculation() 
                if calc_afford and pd.notna(row.get("הון עצמי זמין (₪)")):
                    current_tab.available_funds_entry.delete(0, tk.END)
                    current_tab.available_funds_entry.insert(0, str(int(row["הון עצמי זמין (₪)"])))
                
                for i in range(3):
                    rate_col = f"תרחיש {i+1} - ריבית שנתית (%)"
                    years_col = f"תרחיש {i+1} - שנים להחזר"
                    
                    current_tab.rate_entries[i].delete(0, tk.END)
                    rate_val = row.get(rate_col)
                    if pd.notna(rate_val):
                        current_tab.rate_entries[i].insert(0, str(rate_val))

                    current_tab.years_entries[i].delete(0, tk.END)
                    years_val = row.get(years_col)
                    if pd.notna(years_val):
                        current_tab.years_entries[i].insert(0, str(int(years_val)))
                
                current_tab.calculate() 

            show_error_with_copy("טעינה בוצעה", "הנתונים נטענו בהצלחה מקובץ Excel.", parent=self.root)

        except Exception as e:
            show_error_with_copy("שגיאה בטעינה", f"אירעה שגיאה בעת טעינת הנתונים: {e}", parent=self.root)


if __name__ == "__main__":
    # Configure Matplotlib for Hebrew support
    # Using 'DejaVu Sans' as a fallback if 'Arial Unicode MS' is not available
    plt.rcParams['font.family'] = 'DejaVu Sans' 
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'DejaVu Sans', 'sans-serif'] 
    plt.rcParams['axes.unicode_minus'] = False 

    root = tk.Tk()
    app = MortgageCalculatorApp(root)
    root.mainloop()