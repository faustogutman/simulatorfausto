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
            "חודש": month,
            "קרן": round(principal, 2), 
            "ריבית": round(interest, 2),
            "יתרה": round(max(balance, 0), 2),
            "תשלום חודשי": round(current_monthly_payment, 2)
        })

    return pd.DataFrame(data)

class PropertyTab:
    def __init__(self, parent, idx):
        self.idx = idx
        self.frame = ttk.Frame(parent, padding="10 10 10 10") 

        self.input_frame = ttk.Frame(self.frame)
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

        self.calculate_tab_button = ttk.Button(self.frame, text="חשב נכס זה", command=self.calculate)
        self.calculate_tab_button.pack(pady=10)

        self.results_frame = ttk.Frame(self.frame)
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
            price = 0.0 # Initialize price
            loan_amount = 0.0
            down_payment = 0.0
            purchase_tax = 0.0
            lawyer_fee = 0.0
            broker_fee = 0.0
            total_needed = 0.0

            ltv_str = self.ltv_entry.get()
            if not ltv_str:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "חובה להזין אחוז מימון (LTV).", parent=self.frame)
                return False
            ltv = float(ltv_str)
            if not (0 <= ltv <= 100):
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "אחוז מימון (LTV) חייב להיות בין 0 ל-100.", parent=self.frame)
                return False

            area_str = self.area_entry.get()
            area = float(area_str) if area_str else None
            if area is not None and area <= 0:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "שטח המטר המרובע חייב להיות מספר חיובי.", parent=self.frame)
                return False

            rent_str = self.rent_entry.get()
            rent = float(rent_str) if rent_str else None
            if rent is not None and rent < 0:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "שכירות חודשית צפויה אינה יכולה להיות שלילית.", parent=self.frame)
                return False

            if self.calculate_affordability_var.get():
                available_funds_str = self.available_funds_entry.get()
                if not available_funds_str:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "חובה להזין את סכום הכסף הפנוי.", parent=self.frame)
                    return False
                available_funds = float(available_funds_str)
                if available_funds <= 0:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "סכום הכסף הפנוי חייב להיות חיובי.", parent=self.frame)
                    return False

                # Iterative approach to find the affordable price due to non-linear tax
                estimated_price = available_funds / ((1 - ltv / 100) + LAWYER_FEE_RATE + BROKER_FEE_RATE) # Initial estimate
                
                # Refine the estimate
                tolerance = 1.0 # allowed error in currency (e.g., 1 NIS)
                max_iterations = 100
                current_iteration = 0

                while current_iteration < max_iterations:
                    current_down_payment = estimated_price * (1 - ltv / 100)
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
                        current_total_funds_needed_for_estimated_price = current_down_payment + current_lawyer_fee + current_broker_fee
                    else:
                        current_total_funds_needed_for_estimated_price = current_down_payment + current_purchase_tax + current_lawyer_fee + current_broker_fee

                    diff = available_funds - current_total_funds_needed_for_estimated_price

                    if abs(diff) < tolerance:
                        price = estimated_price
                        break
                    
                    # Adjust estimated price based on the difference
                    estimated_price += diff * 0.5 # Adjust by half the difference to converge

                    current_iteration += 1
                
                if current_iteration == max_iterations:
                    if is_active_tab:
                        messagebox.showwarning("אזהרת חישוב", "לא ניתן למצוא מחיר נכס מדויק עבור ההון העצמי הנתון לאחר מספר רב של ניסיונות. ייתכן שהסכום המחושב הוא קירוב.", parent=self.frame)
                    price = estimated_price # Use the last estimated price as the result

                self.price_entry.config(state='disabled')
                self.price_entry.delete(0, tk.END)
                self.price_entry.insert(0, f"{price:,.0f}")
                self.affordable_price_label.config(text=f"מחיר הנכס המקסימלי שניתן לרכוש: {price:,.0f} ₪")

            else: # If not calculating affordability, take price from input
                price_str = self.price_entry.get()
                if not price_str:
                    if is_active_tab:
                            messagebox.showerror("שגיאת קלט", "חובה להזין מחיר דירה.", parent=self.frame)
                    return False 
                price = float(price_str)
                if price <= 0:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "מחיר הדירה חייב להיות מספר חיובי.", parent=self.frame)
                    return False
                self.price_entry.config(state='normal')
                self.affordable_price_label.config(text="") # Clear label if not calculating affordability

            purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
            
            if self.manual_lawyer_fee_var.get():
                try:
                    lawyer_fee = float(self.lawyer_fee_manual_entry.get())
                    if lawyer_fee < 0:
                        if is_active_tab:
                            messagebox.showerror("שגיאת קלט", "עלות עו\"ד ידנית אינה יכולה להיות שלילית.", parent=self.frame)
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
                            messagebox.showerror("שגיאת קלט", "עלות מתווך ידנית אינה יכולה להיות שלילית.", parent=self.frame)
                        return False
                except ValueError:
                    if is_active_tab:
                        messagebox.showerror("שגיאת קלט", "עלות מתווך ידנית חייבת להיות מספר.", parent=self.frame)
                    return False
            else:
                broker_fee = estimate_broker_fee(price)

            # Recalculate loan_amount and down_payment based on the new checkbox
            base_loan_amount = price * (ltv / 100)
            if self.include_tax_in_mortgage_var.get():
                loan_amount = base_loan_amount + purchase_tax
                down_payment = price - base_loan_amount 
            else:
                loan_amount = base_loan_amount
                down_payment = (price - base_loan_amount) + purchase_tax

            total_needed = down_payment + lawyer_fee + broker_fee

            # If calculating affordability, make sure the final total_needed matches available_funds for display consistency
            if self.calculate_affordability_var.get():
                total_needed = available_funds # Set total_needed to available_funds for display

            self.calculated_results = {
                "purchase_tax": purchase_tax,
                "down_payment": down_payment,
                "loan_amount": loan_amount, 
                "lawyer_fee": lawyer_fee,
                "broker_fee": broker_fee,
                "total_needed": total_needed,
                "price_per_meter": price / area if area is not None and area > 0 else None,
                "rent": rent,
                "input_price": self.price_entry.get(), # Use the actual value from the entry for saving
                "calculated_price": price, # Store the calculated price for internal use
                "input_area": area_str,
                "input_ltv": ltv_str,
                "input_rent": rent_str,
                "input_skip_tax": self.skip_tax_var.get(),
                "input_include_tax_in_mortgage": self.include_tax_in_mortgage_var.get(), # Save the state of the new checkbox
                "input_skip_broker": self.skip_broker_var.get(),
                "input_manual_lawyer_fee": self.manual_lawyer_fee_var.get(), 
                "input_lawyer_fee_manual_value": self.lawyer_fee_manual_entry.get(),
                "input_manual_broker_fee": self.manual_broker_fee_var.get(),
                "input_broker_fee_manual_value": self.broker_fee_manual_entry.get(),
                "input_calculate_affordability": self.calculate_affordability_var.get(),
                "input_available_funds": self.available_funds_entry.get(),
                "input_rates": [], # Initialize lists here to ensure they exist
                "input_years": [], # Initialize lists here to ensure they exist
                "input_alias": self.alias_entry.get(),
                "input_link": self.link_entry.get(),
            }
            self.loan_scenarios_data = [] 
            self.loan_scenarios_rent_comparison = []

            # --- Start of rates and years extraction ---
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
                                messagebox.showerror("שגיאת קלט", f"ריבית שנתית (תרחיש {i+1}) אינה יכולה להיות שלילית.", parent=self.frame)
                            return False
                    except ValueError:
                        if is_active_tab:
                            messagebox.showerror("שגיאת קלט", f"ריבית שנתית (תרחיש {i+1}) חייבת להיות מספר.", parent=self.frame)
                        return False
                    
                    try:
                        current_years = int(years_val)
                        if current_years <= 0:
                            if is_active_tab:
                                messagebox.showerror("שגיאת קלט", f"שנים להחזר (תרחיש {i+1}) חייבות להיות מספר חיובי שלם.", parent=self.frame)
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
                return False # Exit if no valid scenarios to prevent further errors
            
            # Store extracted rates and years in calculated_results
            self.calculated_results["input_rates"] = rates
            self.calculated_results["input_years"] = years
            # --- End of rates and years extraction ---

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
            
            return True 

        except ValueError as e:
            if is_active_tab:
                messagebox.showerror("שגיאת קלט", f"שגיאה בנתונים: {e}\nאנא ודא/י שכל השדות המספריים מולאו נכונה.", parent=self.frame)
            return False
        except Exception as e:
            if is_active_tab:
                messagebox.showerror("שגיאה", f"אירעה שגיאה בלתי צפויה: {e}", parent=self.frame)
            return False


class MortgageApp:
    def __init__(self, root):
        self.root = root
        root.title("מחשבון משכנתא - השוואת נכסים")
        root.geometry("1000x900")
        root.option_add('*font', 'Arial 11')
        root.option_add('*justify', 'right')
        root.tk.call('tk', 'scaling', 1.3) 

        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.main_frame)
        self.v_scroll = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner_frame = tk.Frame(self.canvas)
        
        self.canvas_window_id = self.canvas.create_window(0, 0, window=self.inner_frame, anchor="nw")
        
        self.inner_frame.bind("<Configure>", self._on_inner_frame_configure)
        
        self.canvas.bind("<Configure>", self._center_inner_frame)
        
        self.root.bind("<Configure>", self._on_root_resize)

        self.notebook = ttk.Notebook(self.inner_frame)
        self.notebook.pack(fill="x", expand=True, padx=10, pady=10)

        self.tabs = []
        self.add_tab() # Start with one tab

        self._create_menu()

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="קובץ", menu=file_menu)
        file_menu.add_command(label="הוסף נכס חדש", command=self.add_tab)
        file_menu.add_command(label="שמור נכסים ל-Excel", command=self.save_all_tabs_to_excel)
        file_menu.add_command(label="טען נכסים מ-Excel", command=self.load_tabs_from_excel)
        file_menu.add_separator()
        file_menu.add_command(label="יציאה", command=self.root.quit)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="עזרה", menu=help_menu)
        help_menu.add_command(label="אודות", command=self._show_about)

    def _show_about(self):
        messagebox.showinfo("אודות", "מחשבון משכנתא - השוואת נכסים\nגרסה 1.0\nנוצר על ידי [שם המפתח/חברה]", parent=self.root)

    def _on_inner_frame_configure(self, event):
        # Update the scrollregion of the canvas when the inner frame changes size
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._center_inner_frame()

    def _center_inner_frame(self, event=None):
        # Center the inner_frame horizontally within the canvas
        if self.inner_frame.winfo_width() == 0 or self.canvas.winfo_width() == 0:
            return
        canvas_width = self.canvas.winfo_width()
        frame_width = self.inner_frame.winfo_width()
        
        if frame_width < canvas_width:
            x_offset = (canvas_width - frame_width) // 2
            self.canvas.coords(self.canvas_window_id, x_offset, 0)
        else:
            self.canvas.coords(self.canvas_window_id, 0, 0)

    def _on_root_resize(self, event):
        # When the root window resizes, also re-center the inner frame
        self._center_inner_frame()

    def add_tab(self):
        tab_id = len(self.tabs) + 1
        new_tab = PropertyTab(self.notebook, tab_id)
        self.tabs.append(new_tab)
        self.notebook.add(new_tab.frame, text=f"נכס {tab_id}")
        self.notebook.select(new_tab.frame)

    def save_all_tabs_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return

        all_data = []
        for tab_index, tab in enumerate(self.tabs):
            tab_name = tab.alias_entry.get() if tab.alias_entry.get() else f"נכס {tab.idx}"
            
            # Ensure calculations are up-to-date before saving
            # This will also populate tab.calculated_results and tab.loan_scenarios_data
            tab.calculate() 

            tab_info = {
                "טאב": tab_name,
                "מספר טאב": tab.idx,
                "Alias": tab.alias_entry.get(),
                "Link": tab.link_entry.get(),
                "מחיר דירה (קלט)": tab.calculated_results.get("input_price"),
                "מחיר דירה (מחושב)": tab.calculated_results.get("calculated_price"),
                "מטר מרובע": tab.calculated_results.get("input_area"),
                "אחוז מימון (LTV)": tab.calculated_results.get("input_ltv"),
                "שכירות חודשית צפויה": tab.calculated_results.get("input_rent"),
                "בטל מס רכישה": tab.calculated_results.get("input_skip_tax"),
                "כלול מס רכישה במשכנתא": tab.calculated_results.get("input_include_tax_in_mortgage"), # Save this state
                "הזן עו\"ד ידנית": tab.calculated_results.get("input_manual_lawyer_fee"),
                "עלות עו\"ד ידנית": tab.calculated_results.get("input_lawyer_fee_manual_value"),
                "בטל עלות מתווך": tab.calculated_results.get("input_skip_broker"),
                "הזן מתווך ידנית": tab.calculated_results.get("input_manual_broker_fee"),
                "עלות מתווך ידנית": tab.calculated_results.get("input_broker_fee_manual_value"),
                "חשב מחיר לפי הון עצמי": tab.calculated_results.get("input_calculate_affordability"),
                "הון עצמי זמין": tab.calculated_results.get("input_available_funds"),
                "מס רכישה משוער": f"{tab.calculated_results.get('purchase_tax', 0):,.0f}",
                "הון עצמי נדרש": f"{tab.calculated_results.get('down_payment', 0):,.0f}",
                "סכום הלוואה מהבנק": f"{tab.calculated_results.get('loan_amount', 0):,.0f}",
                "עלות עורך דין משוערת": f"{tab.calculated_results.get('lawyer_fee', 0):,.0f}",
                "עלות מתווך משוערת": f"{tab.calculated_results.get('broker_fee', 0):,.0f}",
                "סה\"כ הון דרוש": f"{tab.calculated_results.get('total_needed', 0):,.0f}",
                "מחיר למטר מרובע": f"{tab.calculated_results.get('price_per_meter', 0):,.2f}" if tab.calculated_results.get('price_per_meter') is not None else "",
            }
            all_data.append(tab_info)

            # Add loan scenarios data
            for i, scenario in enumerate(tab.loan_scenarios_data):
                if scenario: # Only add if scenario data exists
                    scenario_prefix = f"תרחיש {i+1} "
                    scenario_details = {scenario_prefix + k: v for k, v in scenario.items()}
                    tab_info.update(scenario_details)
                    tab_info[f"תרחיש {i+1} השוואת שכירות"] = tab.loan_scenarios_rent_comparison[i]
            
            # Add amortization dataframes as separate sheets or within the same sheet in a structured way
            for i, df in enumerate(tab.df_list):
                if df is not None and not df.empty:
                    # Option 1: Convert DataFrame to a string and store in the main sheet (less ideal for large DFs)
                    # tab_info[f"Amortization Scenario {i+1}"] = df.to_string(index=False)
                    
                    # Option 2: Store DFs separately to be written to separate sheets later
                    # For now, we'll collect them to write after the main summary
                    pass # We'll handle this outside the initial `tab_info` loop

        if not all_data:
            messagebox.showinfo("שמירת נתונים", "אין נתונים לשמירה.", parent=self.root)
            return

        try:
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            
            # Create a DataFrame for the summary of all tabs
            summary_df = pd.DataFrame(all_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Write each amortization schedule to a separate sheet
            for tab_index, tab in enumerate(self.tabs):
                for i, df in enumerate(tab.df_list):
                    if df is not None and not df.empty:
                        sheet_name = f"נכס {tab.idx} תרחיש {i+1}"
                        # Ensure sheet name is not too long
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31] 
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            writer.close()
            messagebox.showinfo("שמירת נתונים", f"הנתונים נשמרו בהצלחה ל:\n{file_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בשמירה", f"אירעה שגיאה בעת שמירת הנתונים: {e}", parent=self.root)

    def load_tabs_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return

        try:
            xls = pd.ExcelFile(file_path)
            summary_df = pd.read_excel(xls, sheet_name='Summary')

            # Clear existing tabs
            for tab_item in self.notebook.tabs():
                self.notebook.forget(tab_item)
            self.tabs = []

            for index, row in summary_df.iterrows():
                self.add_tab() # Add a new tab for each row in the summary
                current_tab = self.tabs[-1] # Get the newly created tab

                # Populate fields based on loaded data
                current_tab.alias_entry.delete(0, tk.END)
                current_tab.alias_entry.insert(0, str(row.get("Alias", "")))
                
                current_tab.link_entry.delete(0, tk.END)
                current_tab.link_entry.insert(0, str(row.get("Link", "")))

                # Handle calculated price vs input price
                if row.get("חשב מחיר לפי הון עצמי", False):
                    current_tab.calculate_affordability_var.set(True)
                    current_tab._toggle_affordability_calculation() # Enable/disable entries
                    current_tab.available_funds_entry.delete(0, tk.END)
                    current_tab.available_funds_entry.insert(0, str(row.get("הון עצמי זמין", "")))
                    # The price_entry will be populated by calculate() based on affordability
                else:
                    current_tab.calculate_affordability_var.set(False)
                    current_tab._toggle_affordability_calculation() # Enable/disable entries
                    current_tab.price_entry.delete(0, tk.END)
                    current_tab.price_entry.insert(0, str(row.get("מחיר דירה (קלט)", "")))

                current_tab.area_entry.delete(0, tk.END)
                current_tab.area_entry.insert(0, str(row.get("מטר מרובע", "")))

                current_tab.ltv_entry.delete(0, tk.END)
                current_tab.ltv_entry.insert(0, str(row.get("אחוז מימון (LTV)", "")))

                current_tab.rent_entry.delete(0, tk.END)
                current_tab.rent_entry.insert(0, str(row.get("שכירות חודשית צפויה", "")))

                current_tab.skip_tax_var.set(bool(row.get("בטל מס רכישה", False)))
                current_tab.include_tax_in_mortgage_var.set(bool(row.get("כלול מס רכישה במשכנתא", False))) # Load this state

                manual_lawyer = bool(row.get("הזן עו\"ד ידנית", False))
                current_tab.manual_lawyer_fee_var.set(manual_lawyer)
                current_tab._toggle_lawyer_fee_entry() # Adjust entry state
                if manual_lawyer:
                    current_tab.lawyer_fee_manual_entry.delete(0, tk.END)
                    current_tab.lawyer_fee_manual_entry.insert(0, str(row.get("עלות עו\"ד ידנית", "")))

                manual_broker = bool(row.get("הזן מתווך ידנית", False))
                current_tab.manual_broker_fee_var.set(manual_broker)
                current_tab._toggle_broker_fee_entry() # Adjust entry state
                if manual_broker:
                    current_tab.broker_fee_manual_entry.delete(0, tk.END)
                    current_tab.broker_fee_manual_entry.insert(0, str(row.get("עלות מתווך ידנית", "")))
                current_tab.skip_broker_var.set(bool(row.get("בטל עלות מתווך", False)))


                for i in range(3):
                    rate_col = f"תרחיש {i+1} ריבית שנתית (%)"
                    years_col = f"תרחיש {i+1} שנים להחזר"
                    
                    rate_val = str(row.get(rate_col, "")).replace(',', '') # Remove commas for float conversion
                    years_val = str(row.get(years_col, "")).replace(',', '')

                    current_tab.rate_entries[i].delete(0, tk.END)
                    current_tab.rate_entries[i].insert(0, rate_val)
                    
                    current_tab.years_entries[i].delete(0, tk.END)
                    current_tab.years_entries[i].insert(0, years_val)
                
                # After loading all input, trigger calculation to display results
                current_tab.calculate()
                
            messagebox.showinfo("טעינת נתונים", "הנתונים נטענו בהצלחה.", parent=self.root)

        except FileNotFoundError:
            messagebox.showerror("שגיאת קובץ", "הקובץ לא נמצא.", parent=self.root)
        except KeyError as e:
            messagebox.showerror("שגיאת פורמט", f"קובץ Excel אינו בפורמט צפוי. חסרה עמודה: {e}", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בטעינה", f"אירעה שגיאה בעת טעינת הנתונים: {e}", parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = MortgageApp(root)
    root.mainloop()