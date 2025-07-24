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


            loan_amount = price * (ltv / 100)
            down_payment = price - loan_amount
            total_needed = down_payment + purchase_tax + lawyer_fee + broker_fee

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
        for i in range(3):
            tab = PropertyTab(self.notebook, i)
            self.tabs.append(tab)
            self.notebook.add(tab.frame, text=f"נכס {i+1}")

        btn_frame = ttk.Frame(self.inner_frame)
        btn_frame.pack(fill="x", pady=10) 

        btn_frame.grid_columnconfigure(0, weight=1) 
        btn_frame.grid_columnconfigure(1, weight=0) 
        btn_frame.grid_columnconfigure(2, weight=0) 
        btn_frame.grid_columnconfigure(3, weight=0) 
        btn_frame.grid_columnconfigure(4, weight=0) 
        btn_frame.grid_columnconfigure(5, weight=0) 
        btn_frame.grid_columnconfigure(6, weight=1) 

        self.save_inputs_button = ttk.Button(btn_frame, text="שמור Inputs ל-CSV", command=self.save_inputs)
        self.save_inputs_button.grid(row=0, column=1, padx=5, pady=5) 

        self.load_inputs_button = ttk.Button(btn_frame, text="טען Inputs מ-CSV", command=self.load_inputs)
        self.load_inputs_button.grid(row=0, column=2, padx=5, pady=5) 

        self.calc_button = ttk.Button(btn_frame, text="חשב עבור כל הנכסים", command=self.calculate_all)
        self.calc_button.grid(row=0, column=3, padx=5, pady=5) 

        self.save_excel_button = ttk.Button(btn_frame, text="שמור נתונים ל-Excel", command=self.save_to_excel)
        self.save_excel_button.grid(row=0, column=4, padx=5, pady=5) 
        
    def _on_inner_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._center_inner_frame() 

    def _on_root_resize(self, event=None):
        self.canvas.update_idletasks() 
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self._center_inner_frame()

    def _center_inner_frame(self, event=None):
        canvas_width = self.canvas.winfo_width()
        inner_frame_width = self.inner_frame.winfo_reqwidth() 
        
        if inner_frame_width > canvas_width:
            self.canvas.coords(self.canvas_window_id, 0, 0) 
        else:
            x_offset = (canvas_width - inner_frame_width) / 2
            self.canvas.coords(self.canvas_window_id, x_offset, 0) 


    def calculate_all(self):
        for tab in self.tabs:
            tab.calculate()
        messagebox.showinfo("חישוב הושלם", "החישובים עבור כל הנכסים הושלמו.", parent=self.root)


    def save_inputs(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="שמור קובץ Inputs"
        )
        if not file_path:
            return
        try:
            data = []
            for tab in self.tabs:
                d = {
                    "Alias": tab.alias_entry.get(),
                    "Link": tab.link_entry.get(),
                    "Price": tab.price_entry.get(),
                    "Area": tab.area_entry.get(),
                    "LTV": tab.ltv_entry.get(),
                    "Rent": tab.rent_entry.get(),
                    "SkipTax": tab.skip_tax_var.get(),
                    "ManualLawyerFee": tab.manual_lawyer_fee_var.get(), 
                    "LawyerFeeManualValue": tab.lawyer_fee_manual_entry.get(), 
                    "ManualBrokerFee": tab.manual_broker_fee_var.get(), 
                    "BrokerFeeManualValue": tab.broker_fee_manual_entry.get(), 
                    "SkipBroker": tab.skip_broker_var.get(),
                    "CalculateAffordability": tab.calculate_affordability_var.get(),
                    "AvailableFunds": tab.available_funds_entry.get(),
                }
                for i in range(3):
                    d[f"Rate{i+1}"] = tab.rate_entries[i].get()
                    d[f"Years{i+1}"] = tab.years_entries[i].get()
                data.append(d)
            df = pd.DataFrame(data)
            df.to_csv(file_path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("הצלחה", "הנתונים נשמרו בהצלחה.", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בשמירה", str(e), parent=self.root)

    def load_inputs(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")],
            title="טען קובץ Inputs"
        )
        if not file_path:
            return
        try:
            df = pd.read_csv(file_path, encoding="utf-8-sig")
            for i, tab in enumerate(self.tabs):
                if i >= len(df):
                    break 
                row = df.iloc[i]
                
                def set_entry_value(entry_widget, value):
                    entry_widget.delete(0, tk.END) 
                    if pd.isna(value): 
                        pass
                    elif isinstance(value, (float, int)):
                        if float(value) == int(float(value)): 
                            entry_widget.insert(0, str(int(float(value))))
                        else:
                            entry_widget.insert(0, str(value))
                    else:
                        entry_widget.insert(0, str(value))

                set_entry_value(tab.alias_entry, row.get("Alias", ""))
                set_entry_value(tab.link_entry, row.get("Link", ""))
                set_entry_value(tab.price_entry, row.get("Price", ""))
                set_entry_value(tab.area_entry, row.get("Area", ""))
                set_entry_value(tab.ltv_entry, row.get("LTV", "70"))
                set_entry_value(tab.rent_entry, row.get("Rent", ""))
                
                tab.skip_tax_var.set(str(row.get("SkipTax", False)).lower() == "true")
                
                manual_lawyer = str(row.get("ManualLawyerFee", False)).lower() == "true"
                tab.manual_lawyer_fee_var.set(manual_lawyer)
                set_entry_value(tab.lawyer_fee_manual_entry, row.get("LawyerFeeManualValue", ""))
                tab._toggle_lawyer_fee_entry() 

                manual_broker = str(row.get("ManualBrokerFee", False)).lower() == "true"
                tab.manual_broker_fee_var.set(manual_broker)
                set_entry_value(tab.broker_fee_manual_entry, row.get("BrokerFeeManualValue", ""))
                tab._toggle_broker_fee_entry() 

                tab.skip_broker_var.set(str(row.get("SkipBroker", False)).lower() == "true")
                if manual_broker: 
                    tab.skip_broker_var.set(False)

                calc_affordability = str(row.get("CalculateAffordability", False)).lower() == "true"
                tab.calculate_affordability_var.set(calc_affordability)
                set_entry_value(tab.available_funds_entry, row.get("AvailableFunds", ""))
                tab._toggle_affordability_calculation() # Call this to set the state of price_entry

                for j in range(3):
                    set_entry_value(tab.rate_entries[j], row.get(f"Rate{j+1}", ""))
                    set_entry_value(tab.years_entries[j], row.get(f"Years{j+1}", ""))
                
                # After loading values into entries, call calculate
                tab.calculate() 
            messagebox.showinfo("הצלחה", "הנתונים נטענו והתעדכנו בהצלחה.", parent=self.root)
        except FileNotFoundError:
            messagebox.showwarning("קובץ לא נמצא", "הקובץ שנבחר לא נמצא.", parent=self.root)
        except pd.errors.EmptyDataError:
            messagebox.showerror("שגיאה בטעינה", "הקובT שנבחר ריק.", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בטעינה", f"אירעה שגיאה בעת טעינת הנתונים: {e}", parent=self.root)

    def save_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="שמור נתוני נכסים ל-Excel"
        )
        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for i, tab in enumerate(self.tabs):
                    calculation_successful = tab.calculate() 

                    alias = tab.alias_entry.get()
                    sheet_name = alias if alias else f"נכס {i+1}"
                    # Sanitize sheet name as Excel sheet names have restrictions
                    sheet_name = sheet_name[:30].replace("/", "_").replace("\\", "_").replace("?", "_").replace("*", "_").replace("[", "_").replace("]", "_").replace(":", "_")
                    
                    if calculation_successful and tab.calculated_results:
                        
                        property_info_data = {
                            "מאפיין": [
                                "כינוי", "לינק", "מחיר דירה (₪)", "מטר מרובע (שטח)",
                                "אחוז מימון (LTV) %", "שכירות חודשית צפויה (₪)",
                                "בטל מס רכישה", 
                                "הזן עלות עו\"ד ידנית", "עלות עו\"ד ידנית (₪)", 
                                "הזן עלות מתווך ידנית", "עלות מתווך ידנית (₪)", 
                                "בטל עלות מתווך",
                                "חשב מחיר נכס לפי הון עצמי", "הון עצמי פנוי (₪)"
                            ],
                            "ערך": [
                                tab.calculated_results.get("input_alias", ""),
                                tab.calculated_results.get("input_link", ""),
                                # Use calculated price if affordability is on, else original input price
                                f"{tab.calculated_results.get('calculated_price', '') if tab.calculated_results.get('input_calculate_affordability') else tab.calculated_results.get('input_price', '')}",
                                tab.calculated_results.get("input_area", ""),
                                tab.calculated_results.get("input_ltv", ""),
                                tab.calculated_results.get("input_rent", ""),
                                "כן" if tab.calculated_results.get("input_skip_tax", False) else "לא",
                                "כן" if tab.calculated_results.get("input_manual_lawyer_fee", False) else "לא", 
                                tab.calculated_results.get("input_lawyer_fee_manual_value", ""), 
                                "כן" if tab.calculated_results.get("input_manual_broker_fee", False) else "לא", 
                                tab.calculated_results.get("input_broker_fee_manual_value", ""), 
                                "כן" if tab.calculated_results.get("input_skip_broker", False) else "לא",
                                "כן" if tab.calculated_results.get("input_calculate_affordability", False) else "לא",
                                tab.calculated_results.get("input_available_funds", ""),
                            ]
                        }
                        property_info_df = pd.DataFrame(property_info_data)
                        property_info_df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False)

                        start_row = len(property_info_df) + 2 

                        key_financials_data = {
                            "מאפיין": [
                                "מס רכישה משוער (₪)",
                                "הון עצמי נדרש (₪)",
                                "סכום הלוואה מהבנק (₪)",
                                "עלות עורך דין משוערת (₪)",
                                "עלות מתווך משוערת (₪)",
                                "סה\"כ הון דרוש (₪)",
                                "מחיר למטר מרובע (₪)"
                            ],
                            "ערך": [
                                f"{tab.calculated_results.get('purchase_tax', 0):,.0f}",
                                f"{tab.calculated_results.get('down_payment', 0):,.0f}",
                                f"{tab.calculated_results.get('loan_amount', 0):,.0f}",
                                f"{tab.calculated_results.get('lawyer_fee', 0):,.0f}",
                                f"{tab.calculated_results.get('broker_fee', 0):,.0f}",
                                f"{tab.calculated_results.get('total_needed', 0):,.0f}",
                                f"{tab.calculated_results.get('price_per_meter', 'N/A'):,.2f}" if tab.calculated_results.get('price_per_meter') is not None else "N/A"
                            ]
                        }
                        key_financials_df = pd.DataFrame(key_financials_data)
                        key_financials_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, startcol=0, index=False)

                        if tab.loan_scenarios_data:
                            start_row += len(key_financials_df) + 2
                            loan_summary_df = pd.DataFrame(tab.loan_scenarios_data)
                            loan_summary_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, startcol=0, index=False)

                            start_row += len(loan_summary_df) + 2
                            rent_comp_df = pd.DataFrame({
                                "השוואת שכירות": tab.loan_scenarios_rent_comparison
                            })
                            rent_comp_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, startcol=0, index=False)

                        current_amort_row = start_row + (len(rent_comp_df) + 2 if tab.loan_scenarios_data else 0) + 2 
                        for df_idx, df in enumerate(tab.df_list): 
                            if df is not None and not df.empty:
                                writer.sheets[sheet_name].cell(row=current_amort_row, column=0, value=f"טבלת פריסה - תרחיש {df_idx+1}")
                                current_amort_row += 1 
                                df.to_excel(writer, sheet_name=sheet_name, startrow=current_amort_row, startcol=0, index=False)
                                current_amort_row += len(df) + 3 

                        img_col_start = 8 
                        for fig_idx, fig in enumerate(tab.figure_list):
                            if tab.df_list[fig_idx] is not None and not tab.df_list[fig_idx].empty: 
                                buf = io.BytesIO()
                                fig.savefig(buf, format='png', bbox_inches='tight')
                                buf.seek(0)
                                img = openpyxl.drawing.image.Image(buf)
                                
                                img.width = 500 
                                img.height = 250 
                                
                                # Calculate anchor row for images to stack correctly
                                if fig_idx == 0:
                                    img.anchor = writer.sheets[sheet_name].cell(row=1, column=img_col_start)
                                else:
                                    # This assumes previous images are added and their height/position is known
                                    # Adjust based on the actual positioning within Excel.
                                    # A more robust way might involve tracking max row used + a fixed offset.
                                    # For simplicity, using a fixed offset for now
                                    img.anchor = writer.sheets[sheet_name].cell(row=writer.sheets[sheet_name]._images[-1].anchor.row + int(writer.sheets[sheet_name]._images[-1].height / 10 + 5), column=img_col_start)
                                    
                                writer.sheets[sheet_name].add_image(img)
                            
                    else:
                        no_data_df = pd.DataFrame({"הודעה": ["אין נתונים זמינים עבור נכס זה (ייתכן שחסרים נתוני קלט או שהחישוב נכשל)."]})
                        no_data_df.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False)
            
            messagebox.showinfo("הצלחה", f"הנתונים נשמרו בהצלחה בקובץ {file_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בשמירה ל-Excel", f"אירעה שגיאה בעת שמירת הקובץ: {e}", parent=self.root)

if __name__ == "__main__":
    root = tk.Tk()
    app = MortgageApp(root)
    root.mainloop()