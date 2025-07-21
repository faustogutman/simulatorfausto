import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd
import io
from PIL import Image # Required for saving images via io.BytesIO
import openpyxl # Explicitly import openpyxl for image handling

# Constants for fees (good practice to define them)
LAWYER_FEE_RATE = 0.015
BROKER_FEE_RATE = 0.02

def calculate_purchase_tax(price):
    # Tax brackets and rates for purchase tax (assuming Israeli tax law for example)
    # This function's logic was confirmed to be correct for progressive tax calculation.
    brackets = [
        (0, 545000, 8),
        (545000, 1362000, 10),
        (1362000, 1890000, 12),
        (1890000, 4890000, 14),
        (4890000, float('inf'), 16),
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
    if monthly_rate == 0:
        return loan_amount / months
    # Handle potential ZeroDivisionError if ((1 + monthly_rate) ** months - 1) is zero
    denominator = ((1 + monthly_rate) ** months - 1)
    if denominator == 0: # This happens if months is 0 or monthly_rate is effectively 0
        return loan_amount / months if months > 0 else 0.0
    return loan_amount * (monthly_rate * (1 + monthly_rate) ** months) / denominator

def generate_amortization_df(loan_amount, annual_rate, years):
    if loan_amount <= 0 or annual_rate < 0 or years <= 0:
        return pd.DataFrame() # Return empty DataFrame for invalid inputs

    monthly_payment = calculate_monthly_payment(loan_amount, annual_rate, years)
    months = years * 12
    balance = loan_amount
    data = []

    for month in range(1, months + 1):
        interest = balance * (annual_rate / 100) / 12
        principal = monthly_payment - interest
        
        # Adjust last payment to precisely zero out the balance
        if balance < principal:
            principal = balance
            # The monthly_payment for the last month should be principal + interest
            # However, the amortization table usually keeps the 'monthly_payment' constant
            # If the last principal payment covers the remaining balance, the balance becomes 0.
            balance = 0
        else:
            balance -= principal
        
        data.append({
            "חודש": month,
            "קרן": round(principal, 2),
            "ריבית": round(interest, 2),
            "יתרה": round(max(balance, 0), 2), # Ensure balance doesn't go negative
            "תשלום חודשי": round(monthly_payment, 2)
        })

    return pd.DataFrame(data)

class PropertyTab:
    def __init__(self, parent, idx):
        self.idx = idx
        self.frame = ttk.Frame(parent, padding="10 10 10 10") # Added padding

        # Use a sub-frame for input fields to allow easy centering later
        self.input_frame = ttk.Frame(self.frame)
        self.input_frame.pack(pady=10) # Pack this frame

        def new_entry():
            return tk.Entry(self.input_frame, justify='right', width=15, font=("Arial", 11))

        padx = 5
        pady = 3
        r = 0 # Row counter for input_frame

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
        self.ltv_entry.insert(0, "70") # Default LTV
        self.ltv_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        ttk.Label(self.input_frame, text="שכירות חודשית צפויה (₪):").grid(row=r, column=0, sticky="e", padx=padx, pady=pady)
        self.rent_entry = new_entry()
        self.rent_entry.grid(row=r, column=1, sticky="w", pady=pady)
        r += 1

        self.skip_tax_var = tk.BooleanVar()
        self.skip_broker_var = tk.BooleanVar()

        self.tax_checkbox = ttk.Checkbutton(self.input_frame, text="בטל מס רכישה", variable=self.skip_tax_var)
        self.tax_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
        r += 1

        self.broker_checkbox = ttk.Checkbutton(self.input_frame, text="בטל עלות מתווך", variable=self.skip_broker_var)
        self.broker_checkbox.grid(row=r, column=0, sticky="w", padx=padx, pady=pady)
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

        # Individual calculate button for this tab
        self.calculate_tab_button = ttk.Button(self.frame, text="חשב נכס זה", command=self.calculate)
        self.calculate_tab_button.pack(pady=10)

        # Frame for results to keep them grouped and allow centering
        self.results_frame = ttk.Frame(self.frame)
        self.results_frame.pack(fill="x", expand=True, pady=10) # Allowed horizontal expansion
        
        # Result labels will be placed within results_frame
        r_res = 0 # Row counter for results_frame
        self.tax_label = ttk.Label(self.results_frame, text="")
        self.tax_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=(10,2))
        r_res += 1

        self.downpayment_label = ttk.Label(self.results_frame, text="")
        self.downpayment_label.grid(row=r_res, column=0, columnspan=2, sticky="w", padx=padx, pady=2)
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

        # Table
        columns = ("loan", "rate", "years", "monthly", "interest", "total")
        self.table = ttk.Treeview(self.results_frame, columns=columns, show="headings", height=6)
        self.table.grid(row=r_res, column=0, columnspan=2, sticky='nsew', pady=10, padx=padx)
        for col, title in zip(columns, ["סכום הלוואה (₪)", "ריבית שנתית (%)", "שנים להחזר", "תשלום חודשי (₪)", "סה\"כ ריבית (₪)", "סה\"כ תשלום כולל (₪)"]):
            self.table.heading(col, text=title)
            self.table.column(col, width=150, anchor="center") # Increased column width
        r_res += 1

        # Configure grid weight for results_frame to make content align nicely
        self.results_frame.grid_rowconfigure(r_res, weight=1)
        self.results_frame.grid_columnconfigure(1, weight=1) # Allow second column (for data) to expand

        # Figures for graphs
        self.figure_list = []
        self.ax_list = []
        self.canvas_list = []
        for i in range(3):
            fig = plt.Figure(figsize=(5, 2.5), dpi=100)
            ax = fig.add_subplot(111)
            canvas = FigureCanvasTkAgg(fig, self.results_frame) # Attach to results_frame
            canvas.get_tk_widget().grid(row=r_res+i, column=0, columnspan=2, pady=3, sticky='nsew')
            self.figure_list.append(fig)
            self.ax_list.append(ax)
            self.canvas_list.append(canvas)

        self.df_list = [None, None, None]

        # Initialize these attributes to avoid AttributeError if not calculated yet
        self.calculated_results = {}
        self.loan_scenarios_data = []
        self.loan_scenarios_rent_comparison = []

    def clear_results(self):
        self.tax_label.config(text="")
        self.downpayment_label.config(text="")
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
        # Ensure these are cleared too when results are cleared
        self.calculated_results = {}
        self.loan_scenarios_data = []
        self.loan_scenarios_rent_comparison = []


    def calculate(self):
        self.clear_results() # Clear previous results before new calculation

        # A flag to check if the current tab is the active one, to avoid multiple messageboxes
        is_active_tab = (self.idx == self.frame.master.index(self.frame)) if hasattr(self.frame.master, 'index') else False


        try:
            price_str = self.price_entry.get()
            if not price_str:
                if is_active_tab:
                     messagebox.showerror("שגיאת קלט", "חובה להזין מחיר דירה.", parent=self.frame)
                return False # Indicate that calculation was not fully successful due to missing price

            price = float(price_str)
            if price <= 0:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "מחיר הדירה חייב להיות מספר חיובי.", parent=self.frame)
                return False

            area_str = self.area_entry.get()
            area = float(area_str) if area_str else None
            if area is not None and area <= 0:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "שטח המטר המרובע חייב להיות מספר חיובי.", parent=self.frame)
                return False

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

            rent_str = self.rent_entry.get()
            rent = float(rent_str) if rent_str else None
            if rent is not None and rent < 0:
                if is_active_tab:
                    messagebox.showerror("שגיאת קלט", "שכירות חודשית צפויה אינה יכולה להיות שלילית.", parent=self.frame)
                return False

            rates = []
            years = []
            valid_scenarios_count = 0
            for i in range(3):
                rate_val = self.rate_entries[i].get()
                years_val = self.years_entries[i].get()
                
                current_rate = None
                current_years = None

                if rate_val and years_val: # Only consider a scenario if both rate and years are provided
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
                # Still proceed to show general property costs, just no loan calculations
        
        except ValueError as e:
            if is_active_tab:
                messagebox.showerror("שגיאת קלט", f"שגיאה בנתונים: {e}\nאנא ודא/י שכל השדות המספריים מולאו נכונה.", parent=self.frame)
            return False
        except Exception as e:
            if is_active_tab:
                messagebox.showerror("שגיאה", f"אירעה שגיאה בלתי צפויה: {e}", parent=self.frame)
            return False

        purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
        lawyer_fee = estimate_lawyer_fee(price)
        broker_fee = 0 if self.skip_broker_var.get() else estimate_broker_fee(price)

        loan_amount = price * (ltv / 100)
        down_payment = price - loan_amount
        total_needed = down_payment + purchase_tax + lawyer_fee + broker_fee

        # Store calculated results as attributes for easy access in save_to_excel
        self.calculated_results = {
            "purchase_tax": purchase_tax,
            "down_payment": down_payment,
            "lawyer_fee": lawyer_fee,
            "broker_fee": broker_fee,
            "total_needed": total_needed,
            "price_per_meter": price / area if area is not None and area > 0 else None,
            "loan_amount": loan_amount,
            "rent": rent,
            "input_price": price_str, # Store raw inputs for Excel
            "input_area": area_str,
            "input_ltv": ltv_str,
            "input_rent": rent_str,
            "input_skip_tax": self.skip_tax_var.get(),
            "input_skip_broker": self.skip_broker_var.get(),
            "input_rates": [r for r in rates], # Store actual rates/years for Excel
            "input_years": [y for y in years],
            "input_alias": self.alias_entry.get(),
            "input_link": self.link_entry.get(),
        }
        self.loan_scenarios_data = [] # To store data for Excel table summary
        self.loan_scenarios_rent_comparison = [] # To store rent comparison strings for Excel

        self.tax_label.config(text=f"מס רכישה משוער: {purchase_tax:,.0f} ₪")
        self.downpayment_label.config(text=f"הון עצמי נדרש: {down_payment:,.0f} ₪")
        self.lawyer_fee_label.config(text=f"עלות עורך דין משוערת: {lawyer_fee:,.0f} ₪")
        self.broker_fee_label.config(text=f"עלות מתווך משוערת: {broker_fee:,.0f} ₪")
        self.total_funds_label.config(text=f"סה\"כ הון דרוש: {total_needed:,.0f} ₪")

        if area is not None and area > 0: # Check if area is provided and valid
            self.price_per_meter_label.config(text=f"מחיר למטר מרובע: {price / area:,.2f} ₪")
        else:
            self.price_per_meter_label.config(text="") # Clear if no area or invalid

        self.table.delete(*self.table.get_children()) # Clear table before filling

        for i in range(3):
            if rates[i] is not None and years[i] is not None:
                df = generate_amortization_df(loan_amount, rates[i], years[i])
                self.df_list[i] = df
                
                if not df.empty:
                    total_interest = df["ריבית"].sum()
                    total_payment = df["תשלום חודשי"].sum()
                    monthly_payment = df["תשלום חודשי"].iloc[0] # Take first monthly payment as base

                    table_row_data = (
                        f"{loan_amount:,.0f}",
                        f"{rates[i]:.2f}",
                        f"{years[i]}",
                        f"{monthly_payment:,.0f}",
                        f"{total_interest:,.0f}",
                        f"{total_payment:,.0f}",
                    )
                    self.table.insert("", "end", values=table_row_data)
                    self.loan_scenarios_data.append({
                        "תרחיש": f"תרחיש {i+1}",
                        "סכום הלוואה (₪)": f"{loan_amount:,.0f}",
                        "ריבית שנתית (%)": f"{rates[i]:.2f}",
                        "שנים להחזר": f"{years[i]}",
                        "תשלום חודשי (₪)": f"{monthly_payment:,.0f}",
                        "סה\"כ ריבית (₪)": f"{total_interest:,.0f}",
                        "סה\"כ תשלום כולל (₪)": f"{total_payment:,.0f}"
                    })


                    rent_compare_str = ""
                    if rent is not None: # Only show if rent is provided and valid
                        ratio = rent / monthly_payment if monthly_payment != 0 else 0
                        rent_compare_str = f"שכירות צפויה: {rent:,.0f} ₪ | תשלום חודשי: {monthly_payment:,.0f} ₪ | יחס שכירות/תשלום: {ratio:.2f}"
                        self.rent_comparison_labels[i].config(text=rent_compare_str)
                    else:
                        self.rent_comparison_labels[i].config(text="") # Clear if rent not provided
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
                    ax.tick_params(axis='both', which='major', labelsize=7) # Smaller ticks
                    self.figure_list[i].tight_layout() # Adjust layout
                    self.canvas_list[i].draw()
                else:
                    self.df_list[i] = None
                    self.table.insert("", "end", values=("אין נתונים עבור תרחיש זה",) * 6)
                    self.rent_comparison_labels[i].config(text="")
                    self.ax_list[i].clear()
                    self.canvas_list[i].draw()
                    self.loan_scenarios_data.append({}) # Append empty dict for this scenario
                    self.loan_scenarios_rent_comparison.append("אין נתוני השוואת שכירות עבור תרחיש זה")
            else:
                self.df_list[i] = None
                self.table.insert("", "end", values=("אין נתונים עבור תרחיש זה (חסר ריבית/שנים)",) * 6)
                self.rent_comparison_labels[i].config(text="")
                self.ax_list[i].clear()
                self.canvas_list[i].draw()
                self.loan_scenarios_data.append({}) # Append empty dict for this scenario
                self.loan_scenarios_rent_comparison.append("אין נתוני השוואת שכירות עבור תרחיש זה")
        
        return True # Indicate successful calculation


class MortgageApp:
    def __init__(self, root):
        self.root = root
        root.title("מחשבון משכנתא - השוואת נכסים")
        root.geometry("1000x900")
        root.option_add('*font', 'Arial 11')
        root.option_add('*justify', 'right')
        root.tk.call('tk', 'scaling', 1.3) # Increased scaling for better readability

        # Main frame to contain canvas and scrollbar
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.main_frame)
        self.v_scroll = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create a frame inside the canvas to hold all content
        # This is the frame we want to center
        self.inner_frame = tk.Frame(self.canvas)
        
        # We need to store the window ID returned by create_window
        self.canvas_window_id = self.canvas.create_window(0, 0, window=self.inner_frame, anchor="nw")
        
        # Bind inner_frame configure to update scrollregion
        self.inner_frame.bind("<Configure>", self._on_inner_frame_configure)
        
        # Bind canvas resize to re-center the inner_frame
        self.canvas.bind("<Configure>", self._center_inner_frame)
        
        # This bind ensures the canvas updates its scrollable area when the root window is resized
        self.root.bind("<Configure>", self._on_root_resize)

        # Tabs
        self.notebook = ttk.Notebook(self.inner_frame)
        self.notebook.pack(fill="x", expand=True, padx=10, pady=10) # Changed to fill="x"

        self.tabs = []
        for i in range(3):
            tab = PropertyTab(self.notebook, i)
            self.tabs.append(tab)
            self.notebook.add(tab.frame, text=f"נכס {i+1}")

        # Action Buttons Frame (for global buttons)
        btn_frame = ttk.Frame(self.inner_frame)
        btn_frame.pack(fill="x", pady=10) 

        # Centering the global buttons using grid
        # We'll use more columns to give finer control and maintain gaps
        btn_frame.grid_columnconfigure(0, weight=1) # Left spacer
        btn_frame.grid_columnconfigure(1, weight=0) # Button 1
        btn_frame.grid_columnconfigure(2, weight=0) # Gap between buttons
        btn_frame.grid_columnconfigure(3, weight=0) # Button 2 (Calculate All)
        btn_frame.grid_columnconfigure(4, weight=0) # Gap between buttons
        btn_frame.grid_columnconfigure(5, weight=0) # Button 3
        btn_frame.grid_columnconfigure(6, weight=1) # Right spacer

        # Place buttons in specific columns
        self.save_inputs_button = ttk.Button(btn_frame, text="שמור Inputs ל-CSV", command=self.save_inputs)
        self.save_inputs_button.grid(row=0, column=1, padx=5, pady=5) # Column 1

        self.load_inputs_button = ttk.Button(btn_frame, text="טען Inputs מ-CSV", command=self.load_inputs)
        self.load_inputs_button.grid(row=0, column=2, padx=5, pady=5) # Column 2

        self.calc_button = ttk.Button(btn_frame, text="חשב עבור כל הנכסים", command=self.calculate_all)
        self.calc_button.grid(row=0, column=3, padx=5, pady=5) # Column 3 (Center)

        self.save_excel_button = ttk.Button(btn_frame, text="שמור נתונים ל-Excel", command=self.save_to_excel)
        self.save_excel_button.grid(row=0, column=4, padx=5, pady=5) # Column 4
        
    def _on_inner_frame_configure(self, event=None):
        # Update canvas scroll region when the inner_frame's size changes
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._center_inner_frame() # Also try to re-center on content change

    def _on_root_resize(self, event=None):
        # Update canvas scroll region and re-center inner_frame when the root window is resized
        self.canvas.update_idletasks() # Ensure all widgets are rendered
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self._center_inner_frame()

    def _center_inner_frame(self, event=None):
        # Center the inner_frame within the canvas
        canvas_width = self.canvas.winfo_width()
        inner_frame_width = self.inner_frame.winfo_reqwidth() # Requested width
        
        # If inner_frame is wider than canvas, don't center horizontally, align left
        if inner_frame_width > canvas_width:
            self.canvas.coords(self.canvas_window_id, 0, 0) # Align to left (nw anchor)
        else:
            x_offset = (canvas_width - inner_frame_width) / 2
            self.canvas.coords(self.canvas_window_id, x_offset, 0) # Center horizontally, align top (nw anchor initially)


    def calculate_all(self):
        # This function should call calculate on each tab without showing multiple popups
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
                    "SkipBroker": tab.skip_broker_var.get(),
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
                    break # Stop if CSV has fewer rows than tabs
                row = df.iloc[i]
                
                # Helper to set entry value, converting floats like X.0 to int X
                def set_entry_value(entry_widget, value):
                    entry_widget.delete(0, tk.END) # Clear existing content first
                    if pd.isna(value): # Check for NaN from pandas
                        # Leave empty
                        pass
                    elif isinstance(value, (float, int)):
                        if float(value) == int(float(value)): # Check if float is a whole number
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
                
                # Handle booleans carefully, accounting for string "True"/"False" from CSV
                tab.skip_tax_var.set(str(row.get("SkipTax", False)).lower() == "true")
                tab.skip_broker_var.set(str(row.get("SkipBroker", False)).lower() == "true")
                
                for j in range(3):
                    set_entry_value(tab.rate_entries[j], row.get(f"Rate{j+1}", ""))
                    set_entry_value(tab.years_entries[j], row.get(f"Years{j+1}", ""))
                
                # Clear and recalculate each tab after loading its data
                # No need to clear_results here as calculate() does it internally
                tab.calculate() # Re-calculate immediately after loading
            messagebox.showinfo("הצלחה", "הנתונים נטענו והתעדכנו בהצלחה.", parent=self.root)
        except FileNotFoundError:
            messagebox.showwarning("קובץ לא נמצא", "הקובץ שנבחר לא נמצא.", parent=self.root)
        except pd.errors.EmptyDataError:
            messagebox.showerror("שגיאה בטעינה", "הקובץ שנבחר ריק.", parent=self.root)
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
                    # Always try to calculate the tab to ensure data is fresh
                    # The calculate method now returns False if essential inputs are missing/invalid
                    calculation_successful = tab.calculate() 

                    alias = tab.alias_entry.get()
                    sheet_name = alias if alias else f"נכס {i+1}"
                    # Sanitize sheet name as Excel sheet names have limits
                    sheet_name = sheet_name[:30].replace("/", "_").replace("\\", "_").replace("?", "_").replace("*", "_").replace("[", "_").replace("]", "_").replace(":", "_")
                    
                    if not calculation_successful or not tab.calculated_results:
                        # Create an empty sheet or a sheet with a message if calculation failed
                        dummy_df = pd.DataFrame([{"הערה": "אין נתונים לחישוב עבור נכס זה או שחסרים קלטים חיוניים."}])
                        dummy_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        continue # Skip to the next tab if calculation failed

                    # --- Write Inputs and Summary Results ---
                    inputs_summary_data = {
                        "פרט": ["כינוי", "לינק", "מחיר דירה", "מטר מרובע", "אחוז מימון (LTV)", "שכירות חודשית צפויה", "בטל מס רכישה", "בטל עלות מתווך",
                                "מס רכישה משוער", "הון עצמי נדרש", "עלות עו\"ד משוערת", "עלות מתווך משוערת", "סה\"כ הון דרוש", "מחיר למטר מרובע", "סכום הלוואה"],
                        "ערך": [
                            tab.calculated_results.get("input_alias", ""),
                            tab.calculated_results.get("input_link", ""),
                            f"{float(tab.calculated_results['input_price']):,.0f} ₪" if tab.calculated_results.get("input_price") else "",
                            f"{float(tab.calculated_results['input_area']):,.0f}" if tab.calculated_results.get("input_area") else "",
                            f"{float(tab.calculated_results['input_ltv']):.0f}%" if tab.calculated_results.get("input_ltv") else "",
                            f"{float(tab.calculated_results['input_rent']):,.0f} ₪" if tab.calculated_results.get("input_rent") else "",
                            "כן" if tab.calculated_results.get("input_skip_tax", False) else "לא",
                            "כן" if tab.calculated_results.get("input_skip_broker", False) else "לא",
                            f"{tab.calculated_results['purchase_tax']:,.0f} ₪",
                            f"{tab.calculated_results['down_payment']:,.0f} ₪",
                            f"{tab.calculated_results['lawyer_fee']:,.0f} ₪",
                            f"{tab.calculated_results['broker_fee']:,.0f} ₪",
                            f"{tab.calculated_results['total_needed']:,.0f} ₪",
                            f"{tab.calculated_results['price_per_meter']:,.2f} ₪" if tab.calculated_results['price_per_meter'] is not None else "N/A",
                            f"{tab.calculated_results['loan_amount']:,.0f} ₪"
                        ]
                    }
                    inputs_summary_df = pd.DataFrame(inputs_summary_data)
                    
                    # Write inputs and summary to Excel
                    inputs_summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0, startcol=0)
                    
                    # Get the worksheet object to write images
                    worksheet = writer.sheets[sheet_name]
                    
                    # Start row for loan scenarios
                    current_row = inputs_summary_df.shape[0] + 3 # Adjust position after inputs/summary

                    # --- Write All Loan Scenarios ---
                    for j in range(3):
                        scenario_alias = f"תרחיש {j+1}"
                        
                        rate_val = tab.calculated_results["input_rates"][j] if j < len(tab.calculated_results["input_rates"]) else None
                        years_val = tab.calculated_results["input_years"][j] if j < len(tab.calculated_results["input_years"]) else None

                        worksheet.cell(row=current_row + 1, column=1, value=f"--- {scenario_alias} (ריבית: {rate_val if rate_val is not None else 'N/A'}%, שנים: {years_val if years_val is not None else 'N/A'}) ---")
                        current_row += 2 # Move past the header
                        
                        if tab.df_list[j] is not None and not tab.df_list[j].empty:
                            tab.df_list[j].to_excel(writer, sheet_name=sheet_name, index=False, startrow=current_row, startcol=0)
                            current_row += tab.df_list[j].shape[0] + 2 # Move past the table and add a gap

                            # Add the rent comparison string
                            if j < len(tab.loan_scenarios_rent_comparison):
                                worksheet.cell(row=current_row, column=1, value=tab.loan_scenarios_rent_comparison[j])
                            current_row += 2 # Move past rent comparison string

                            # Save and embed the graph
                            fig = tab.figure_list[j]
                            buf = io.BytesIO()
                            fig.savefig(buf, format='png', bbox_inches='tight') # Save to buffer
                            buf.seek(0)
                            
                            img = Image.open(buf)
                            image_path = io.BytesIO()
                            img.save(image_path, format='png')
                            image_path.seek(0) # Rewind to start of buffer for reading

                            worksheet.add_image(openpyxl.drawing.image.Image(image_path), f'A{current_row}')
                            # Estimate rows needed for image based on its height/dpi. Adjust multiplier if needed.
                            # Standard row height is about 15 pixels. Image height in pixels / 15.
                            current_row += int(fig.get_figheight() * fig.dpi / 15) + 2
                            
                        else:
                            worksheet.cell(row=current_row + 1, column=1, value=f"אין נתוני תשלום זמינים עבור תרחיש זה (חסר ריבית/שנים).")
                            current_row += 3 # Just a gap for empty scenarios

            messagebox.showinfo("הצלחה", f"כל נתוני הנכסים נשמרו בהצלחה ל-Excel:\n{file_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("שגיאה בשמירה", f"אירעה שגיאה בעת שמירה ל-Excel: {e}", parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = MortgageApp(root)
    root.mainloop()