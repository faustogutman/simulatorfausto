import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd
import io
from PIL import Image
import openpyxl

# Constants for fees (good practice to define them)
LAWYER_FEE_RATE = 0.01
BROKER_FEE_RATE = 0.02

def calculate_purchase_tax(price):
    # Tax brackets and rates for purchase tax (assuming Israeli tax law for example)
    brackets = [
        (0, 6055070, 8),
        (6055070, float('inf'), 10),
    ]
    tax = 0
    remaining_price = price
    for low, high, rate in brackets:
        if remaining_price > low:
            # Calculate the portion of the remaining price that falls into the current bracket
            taxable_in_current_bracket = min(high, remaining_price) - low if remaining_price > low else 0
            tax += taxable_in_current_bracket * rate / 100
            remaining_price -= taxable_in_current_bracket # Deduct the portion that was just taxed
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
    if monthly_rate == 0: # Avoid division by zero if rate is 0
        return loan_amount / months if months > 0 else 0.0
    
    if abs(monthly_rate) < 1e-9: # If rate is extremely close to zero
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
        if balance <= 0: # Loan paid off
            break

        interest = balance * (annual_rate / 100) / 12
        principal = current_monthly_payment - interest
        
        principal = min(principal, balance) # Ensure principal doesn't exceed balance
        
        balance -= principal
        
        data.append({
            "חודש": round(month, 2), # Round month for consistency
            "קרן": round(principal, 2), 
            "ריבית": round(interest, 2),
            "יתרה": round(max(balance, 0), 2),
            "תשלום חודשי": round(current_monthly_payment, 2)
        })

    return pd.DataFrame(data)

class PropertyTab:
    def __init__(self, parent, idx):
        self.idx = idx
        self.frame = ttk.Frame(parent) # This frame will contain the canvas and scrollbar
        self.frame.pack(expand=True, fill="both") # Make sure the main frame expands

        # Create a Canvas widget
        self.canvas = tk.Canvas(self.frame, borderwidth=0, background="#f0f0f0") # Add background for visibility
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create a Scrollbar and link it to the Canvas
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")

        # Configure the Canvas to use the scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind('<Configure>', self._on_canvas_configure) # Recalculate scroll region on resize
        self.canvas.bind_all('<MouseWheel>', self._on_mousewheel) # Bind mouse wheel for scrolling

        # Create a frame to hold all the tab's content inside the canvas
        self.content_frame = ttk.Frame(self.canvas, padding="10 10 10 10")
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # --- Now, all your existing widgets will be packed into self.content_frame ---

        self.input_frame = ttk.Frame(self.content_frame) # Change parent to content_frame
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

        # New: Checkbox to include tax in mortgage
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
        
        # New: Calculate affordability checkbox and entry
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

        self.calculate_tab_button = ttk.Button(self.content_frame, text="חשב נכס זה", command=self.calculate) # Change parent to content_frame
        self.calculate_tab_button.pack(pady=10)

        self.results_frame = ttk.Frame(self.content_frame) # Change parent to content_frame
        self.results_frame.pack(fill="x", expand=True, pady=10) 
        
        r_res = 0 

        # New: Affordable Price Label
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

        # Ensure the canvas's scroll region is updated initially and whenever content changes
        self.content_frame.bind('<Configure>', self._on_frame_configure)

    def _on_frame_configure(self, event=None):
        """Update the scrollregion of the canvas based on the content frame size."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        """Adjust content frame width to match canvas width when canvas resizes."""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw"), width=canvas_width)


    def _on_mousewheel(self, event):
        """Enable mouse wheel scrolling."""
        # For Windows/Linux: event.delta is typically +/-120 per scroll "click"
        # For macOS: event.delta is typically a smaller value, actual pixels
        if event.delta: # Check if delta exists (it does for mousewheel)
            # Normalize delta to be consistent across platforms
            if event.num == 5 or event.delta < 0: # Mouse wheel down
                self.canvas.yview_scroll(1, "unit")
            elif event.num == 4 or event.delta > 0: # Mouse wheel up
                self.canvas.yview_scroll(-1, "unit")


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
            self.price_entry.config(state='disabled') # Disable price entry if calculating affordability
        else:
            self.available_funds_entry.config(state='disabled')
            self.available_funds_entry.delete(0, tk.END)
            self.price_entry.config(state='normal') # Re-enable price entry

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
            # Initialize core financial values
            price = 0.0
            loan_amount = 0.0
            down_payment = 0.0
            purchase_tax = 0.0
            lawyer_fee = 0.0
            broker_fee = 0.0
            total_needed = 0.0

            # --- Input Validation and Extraction ---
            ltv_str = self.ltv_entry.get()
            if not ltv_str:
                if is_active_tab:
                    messagebox.showerror("קלט חסר", "יש להזין אחוז מימון (LTV).", parent=self.frame)
                return False
            ltv = float(ltv_str)
            if not (0 <= ltv <= 100):
                if is_active_tab:
                    messagebox.showerror("קלט לא חוקי", "אחוז מימון (LTV) חייב להיות בין 0 ל-100.", parent=self.frame)
                return False

            area_str = self.area_entry.get()
            area = float(area_str) if area_str else None
            if area is not None and area <= 0:
                if is_active_tab:
                    messagebox.showerror("קלט לא חוקי", "שטח המטר המרובע חייב להיות מספר חיובי.", parent=self.frame)
                return False

            rent_str = self.rent_entry.get()
            rent = float(rent_str) if rent_str else None
            if rent is not None and rent < 0:
                if is_active_tab:
                    messagebox.showerror("קלט לא חוקי", "שכירות חודשית צפויה אינה יכולה להיות שלילית.", parent=self.frame)
                return False

            # --- Affordability Calculation Logic ---
            if self.calculate_affordability_var.get():
                available_funds_str = self.available_funds_entry.get()
                if not available_funds_str:
                    if is_active_tab:
                        messagebox.showerror("קלט חסר", "יש להזין את סכום הכסף הפנוי.", parent=self.frame)
                    return False
                available_funds = float(available_funds_str)
                if available_funds <= 0:
                    if is_active_tab:
                        messagebox.showerror("קלט לא חוקי", "סכום הכסף הפנוי חייב להיות חיובי.", parent=self.frame)
                    return False

                # Iterative approach to find the affordable price due to non-linear tax
                estimated_price = available_funds / ((1 - ltv / 100) + LAWYER_FEE_RATE + BROKER_FEE_RATE) # Initial guess
                
                tolerance = 1.0 # Tolerance for price convergence (e.g., 1 NIS)
                max_iterations = 100
                current_iteration = 0

                while current_iteration < max_iterations:
                    current_down_payment_ratio = (1 - ltv / 100) # Percentage of price for down payment
                    current_down_payment_from_price = estimated_price * current_down_payment_ratio
                    current_purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(estimated_price)
                    
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

                    # Calculate total funds needed based on whether tax is included in mortgage
                    if self.include_tax_in_mortgage_var.get():
                        # If tax is included in the mortgage, it's part of the loan, not the direct cash outlay.
                        # So, available funds cover down payment (price - loan), lawyer, and broker.
                        # The loan would be (estimated_price * LTV) + current_purchase_tax.
                        # Down payment is estimated_price - loan.
                        # So, total cash needed = (estimated_price - ((estimated_price * LTV) + current_purchase_tax)) + lawyer + broker
                        # Simplified: estimated_price * (1 - LTV/100) - current_purchase_tax + lawyer + broker
                        current_loan_amount_for_affordability = (estimated_price * (ltv / 100)) + current_purchase_tax
                        current_down_payment_for_affordability = estimated_price - current_loan_amount_for_affordability
                        current_total_funds_needed = current_down_payment_for_affordability + current_lawyer_fee + current_broker_fee
                    else:
                        # If tax is NOT included in the mortgage, it's part of the cash outlay.
                        current_total_funds_needed = current_down_payment_from_price + current_purchase_tax + current_lawyer_fee + current_broker_fee

                    diff = available_funds - current_total_funds_needed

                    if abs(diff) < tolerance:
                        price = estimated_price
                        break
                    
                    # Adjust estimated price based on the difference for convergence
                    # Using a smaller factor (e.g., 0.1 to 0.5) can help prevent overshooting,
                    # especially with non-linear tax.
                    estimated_price += diff * 0.5 

                    current_iteration += 1
                
                if current_iteration == max_iterations:
                    if is_active_tab:
                        messagebox.showwarning("אזהרת חישוב", "לא ניתן למצוא מחיר נכס מדויק עבור ההון העצמי הנתון לאחר מספר רב של ניסיונות. ייתכן שהסכום המחושב הוא קירוב.", parent=self.frame)
                    price = estimated_price # Use the last estimated price as the result

                self.price_entry.config(state='disabled')
                self.price_entry.delete(0, tk.END)
                self.price_entry.insert(0, f"{price:,.0f}")
                self.affordable_price_label.config(text=f"מחיר הנכס המקסימלי שניתן לרכוש: {price:,.0f} ₪")

            else: # If not calculating affordability, get price from input
                price_str = self.price_entry.get()
                if not price_str:
                    if is_active_tab:
                        messagebox.showerror("קלט חסר", "יש להזין מחיר דירה.", parent=self.frame)
                    return False 
                price = float(price_str)
                if price <= 0:
                    if is_active_tab:
                        messagebox.showerror("קלט לא חוקי", "מחיר הדירה חייב להיות מספר חיובי.", parent=self.frame)
                    return False
                self.price_entry.config(state='normal')
                self.affordable_price_label.config(text="") # Clear label if not calculating affordability

            # --- Calculate Fees and Loan Amounts ---
            purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
            
            if self.manual_lawyer_fee_var.get():
                try:
                    lawyer_fee = float(self.lawyer_fee_manual_entry.get())
                    if lawyer_fee < 0:
                        if is_active_tab:
                            messagebox.showerror("קלט לא חוקי", "עלות עו\"ד ידנית אינה יכולה להיות שלילית.", parent=self.frame)
                        return False
                except ValueError:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "עלות עו\"ד ידנית חייבת להיות מספר.", parent=self.frame)
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
                            messagebox.showerror("קלט לא חוקי", "עלות מתווך ידנית אינה יכולה להיות שלילית.", parent=self.frame)
                        return False
                except ValueError:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "עלות מתווך ידנית חייבת להיות מספר.", parent=self.frame)
                    return False
            else:
                broker_fee = estimate_broker_fee(price)

            # Recalculate loan_amount and down_payment based on tax inclusion
            base_loan_amount = price * (ltv / 100)
            if self.include_tax_in_mortgage_var.get():
                loan_amount = base_loan_amount + purchase_tax
                # If tax is included in the mortgage, the down payment is simply the price minus the *new, higher* loan amount.
                down_payment = price - loan_amount 
            else:
                loan_amount = base_loan_amount
                # If tax is NOT included, down payment covers the property's non-financed portion PLUS the tax.
                down_payment = (price - base_loan_amount) + purchase_tax

            total_needed = down_payment + lawyer_fee + broker_fee

            # If calculating affordability, ensure the final total_needed matches available_funds for display consistency
            if self.calculate_affordability_var.get():
                total_needed = available_funds # Set total_needed to available_funds for display

            # Store calculated results
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

            # --- Extract and Validate Loan Scenarios (Rates and Years) ---
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
                                messagebox.showerror("קלט לא חוקי", f"ריבית שנתית (תרחיש {i+1}) אינה יכולה להיות שלילית.", parent=self.frame)
                            return False
                    except ValueError:
                        if is_active_tab:
                            messagebox.showerror("שגיאת קלט", f"ריבית שנתית (תרחיש {i+1}) חייבת להיות מספר.", parent=self.frame)
                        return False
                    
                    try:
                        current_years = int(years_val)
                        if current_years <= 0:
                            if is_active_tab:
                                messagebox.showerror("קלט לא חוקי", f"שנים להחזר (תרחיש {i+1}) חייבות להיות מספר חיובי שלם.", parent=self.frame)
                            return False
                    except ValueError:
                        if is_active_tab:
                            messagebox.showerror("שגיאת קלט", f"שנים להחזר (תרחיש {i+1}) חייבות להיות מספר שלם.", parent=self.frame)
                        return False
                    valid_scenarios_count += 1
                
                rates.append(current_rate)
                years.append(current_years)

            if valid_scenarios_count == 0:
                if is_active_tab:
                    messagebox.showwarning("אין נתונים לחישוב", "אנא הזן/י לפחות ריבית שנתית אחת ושנים להחזר עבור תרחיש.", parent=self.frame)
                return False 
            
            self.calculated_results["input_rates"] = rates
            self.calculated_results["input_years"] = years

            # --- Display Results ---
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

            # --- Process and Display Loan Scenarios ---
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
            
            # After updating all content, re-configure scroll region
            self._on_frame_configure()
            
            return True 

        except ValueError as e:
            if is_active_tab:
                messagebox.showerror("שגיאת קלט", f"שגיאה בנתונים: {e}\nאנא ודא/י שכל השדות המספריים מולאו נכונה.", parent=self.frame)
            return False
        except Exception as e:
            if is_active_tab:
                messagebox.showerror("שגיאה כללית", f"אירעה שגיאה בלתי צפויה: {e}", parent=self.frame)
            return False

# --- Main Application Window ---
class MortgageCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("מחשבון נדל\"ן מקיף")
        self.root.geometry("1200x900") # Increased default size for better visibility

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.property_tabs = []
        self.add_tab() # Start with one tab

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
        new_tab = PropertyTab(self.notebook, idx)
        self.property_tabs.append(new_tab)
        self.notebook.add(new_tab.frame, text=f"נכס {idx + 1}")
        self.notebook.select(new_tab.frame) # Switch to the new tab

    def save_data(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", # Corrected here
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="שמור נתוני נכסים")
        if not filepath:
            return

        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                summary_data = []
                for idx, prop_tab in enumerate(self.property_tabs):
                    prop_tab.calculate() # This populates calculated_results and loan_scenarios_data
                    
                    results = prop_tab.calculated_results
                    loan_scenarios = prop_tab.loan_scenarios_data
                    rent_comparisons = prop_tab.loan_scenarios_rent_comparison

                    alias = results.get("input_alias", f"נכס {idx + 1}")
                    link = results.get("input_link", "")

                    # Main property details
                    summary_row = {
                        "Alias": alias,
                        "Link": link,
                        "מחיר דירה (₪)": results.get("calculated_price"),
                        "מטר מרובע (שטח)": results.get("input_area"),
                        "אחוז מימון (LTV) %": results.get("input_ltv"),
                        "שכירות חודשית צפויה (₪)": results.get("rent"),
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
                    
                    # Add loan scenarios to summary row
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

                    # Write amortization tables to separate sheets if available
                    for i, df in enumerate(prop_tab.df_list):
                        if df is not None and not df.empty:
                            sheet_name = f"{alias}_תרחיש_{i+1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)

                pd.DataFrame(summary_data).to_excel(writer, sheet_name="סיכום נכסים", index=False)
            
            messagebox.showinfo("שמירה בוצעה", "הנתונים נשמרו בהצלחה לקובץ Excel.", parent=self.root)

        except Exception as e:
            messagebox.showerror("שגיאה בשמירה", f"אירעה שגיאה בעת שמירת הנתונים: {e}", parent=self.root)

    def load_data(self):
        filepath = filedialog.askopenfilename(defaultextension=".xlsx", # Corrected here
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="טען נתוני נכסים")
        if not filepath:
            return

        try:
            xls = pd.ExcelFile(filepath)
            
            # Clear existing tabs
            for _ in range(len(self.property_tabs)):
                self.notebook.forget(0)
            self.property_tabs = []

            if "סיכום נכסים" not in xls.sheet_names:
                messagebox.showerror("שגיאה בטעינה", "קובץ Excel אינו מכיל גיליון 'סיכום נכסים'.", parent=self.root)
                return

            summary_df = pd.read_excel(xls, sheet_name="סיכום נכסים")

            for index, row in summary_df.iterrows():
                self.add_tab()
                current_tab = self.property_tabs[-1]

                # Populate input fields
                current_tab.alias_entry.delete(0, tk.END)
                current_tab.alias_entry.insert(0, row.get("Alias", ""))
                
                current_tab.link_entry.delete(0, tk.END)
                current_tab.link_entry.insert(0, row.get("Link", ""))

                current_tab.price_entry.delete(0, tk.END)
                if not row.get("חשב מחיר נכס לפי הון עצמי"): # Only insert if not affordability calculation
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

                # Manual lawyer fee
                manual_lawyer = (row.get("הזן עלות עו\"ד ידנית") == "כן")
                current_tab.manual_lawyer_fee_var.set(manual_lawyer)
                current_tab._toggle_lawyer_fee_entry() # Update state of entry
                if manual_lawyer and pd.notna(row.get("עלות עו\"ד ידנית")):
                    current_tab.lawyer_fee_manual_entry.delete(0, tk.END)
                    current_tab.lawyer_fee_manual_entry.insert(0, str(int(row["עלות עו\"ד ידנית"])))

                # Broker fee
                manual_broker = (row.get("הזן עלות מתווך ידנית") == "כן")
                current_tab.manual_broker_fee_var.set(manual_broker)
                current_tab._toggle_broker_fee_entry() # Update state of entry
                if manual_broker and pd.notna(row.get("עלות מתווך ידנית")):
                    current_tab.broker_fee_manual_entry.delete(0, tk.END)
                    current_tab.broker_fee_manual_entry.insert(0, str(int(row["עלות מתווך ידנית"])))
                current_tab.skip_broker_var.set(row.get("בטל עלות מתווך") == "כן")

                # Affordability calculation
                calc_afford = (row.get("חשב מחיר נכס לפי הון עצמי") == "כן")
                current_tab.calculate_affordability_var.set(calc_afford)
                current_tab._toggle_affordability_calculation() # Update state of entry
                if calc_afford and pd.notna(row.get("הון עצמי זמין (₪)")):
                    current_tab.available_funds_entry.delete(0, tk.END)
                    current_tab.available_funds_entry.insert(0, str(int(row["הון עצמי זמין (₪)"])))
                
                # Loan scenarios (rates and years)
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
                
                # Re-run calculation for each loaded tab to populate results and graphs
                current_tab.calculate() 

            messagebox.showinfo("טעינה בוצעה", "הנתונים נטענו בהצלחה מקובץ Excel.", parent=self.root)

        except Exception as e:
            messagebox.showerror("שגיאה בטעינה", f"אירעה שגיאה בעת טעינת הנתונים: {e}", parent=self.root)


if __name__ == "__main__":
    # Configure Matplotlib for Hebrew support
    plt.rcParams['font.family'] = 'DejaVu Sans' # A common font that supports Hebrew
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'sans-serif'] # Fallback fonts
    plt.rcParams['axes.unicode_minus'] = False # This is important to display minus signs correctly

    root = tk.Tk()
    app = MortgageCalculatorApp(root)
    root.mainloop()