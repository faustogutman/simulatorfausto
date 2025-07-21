import tkinter as tk
from tkinter import ttk, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd

def calculate_purchase_tax(price):
    brackets = [
        (0, 545000, 8),
        (545000, 1362000, 10),
        (1362000, 1890000, 12),
        (1890000, 4890000, 14),
        (4890000, float('inf'), 16),
    ]
    tax = 0
    remaining = price
    for low, high, rate in brackets:
        if price > low:
            taxable = min(high - low, remaining)
            tax += taxable * rate / 100
            remaining -= taxable
            if remaining <= 0:
                break
    return tax

def estimate_lawyer_fee(price):
    return price * 0.015  # 1.5%

def estimate_broker_fee(price):
    return price * 0.02  # 2%

def calculate_monthly_payment(loan_amount, annual_rate, years):
    months = years * 12
    monthly_rate = (annual_rate / 100) / 12
    if monthly_rate == 0:
        return loan_amount / months
    return loan_amount * (monthly_rate * (1 + monthly_rate) ** months) / ((1 + monthly_rate) ** months - 1)

def generate_amortization_df(loan_amount, annual_rate, years):
    monthly_payment = calculate_monthly_payment(loan_amount, annual_rate, years)
    months = years * 12
    balance = loan_amount
    data = []

    for month in range(1, months + 1):
        interest = balance * (annual_rate / 100) / 12
        principal = monthly_payment - interest
        balance -= principal
        data.append({
            "חודש": month,
            "קרן": round(principal, 2),
            "ריבית": round(interest, 2),
            "יתרה": round(max(balance, 0), 2),
            "תשלום חודשי": round(monthly_payment, 2)
        })

    return pd.DataFrame(data)

class MortgageApp:
    def __init__(self, root):
        self.root = root
        root.title("מחשבון משכנתא")

        self.canvas = tk.Canvas(root, borderwidth=0)
        self.frame = tk.Frame(self.canvas)
        self.vsb = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw")

        self.frame.bind("<Configure>", self.on_frame_configure)

        row_idx = 0
        tk.Label(self.frame, text="מחיר דירה (₪):").grid(row=row_idx, column=0, sticky="e")
        self.price_entry = tk.Entry(self.frame)
        self.price_entry.grid(row=row_idx, column=1)
        row_idx += 1

        tk.Label(self.frame, text="מטר מרובע (שטח):").grid(row=row_idx, column=0, sticky="e")
        self.area_entry = tk.Entry(self.frame)
        self.area_entry.grid(row=row_idx, column=1)
        row_idx += 1

        tk.Label(self.frame, text="אחוז מימון (LTV) %:").grid(row=row_idx, column=0, sticky="e")
        self.ltv_entry = tk.Entry(self.frame)
        self.ltv_entry.insert(0, "70")
        self.ltv_entry.grid(row=row_idx, column=1)
        row_idx += 1

        # Checkboxes לביטול מס רכישה ומתווך
        self.skip_tax_var = tk.BooleanVar()
        self.skip_broker_var = tk.BooleanVar()

        self.tax_checkbox = tk.Checkbutton(self.frame, text="בטל מס רכישה", variable=self.skip_tax_var, command=self.calculate)
        self.tax_checkbox.grid(row=row_idx, column=0, sticky="w", pady=2)
        row_idx += 1

        self.broker_checkbox = tk.Checkbutton(self.frame, text="בטל עלות מתווך", variable=self.skip_broker_var, command=self.calculate)
        self.broker_checkbox.grid(row=row_idx, column=0, sticky="w", pady=2)
        row_idx += 1

        self.rate_entries = []
        self.years_entries = []
        for i in range(3):
            tk.Label(self.frame, text=f"ריבית שנתית (תרחיש {i+1}) %:").grid(row=row_idx, column=0, sticky="e")
            rate_entry = tk.Entry(self.frame)
            rate_entry.grid(row=row_idx, column=1)
            self.rate_entries.append(rate_entry)
            row_idx += 1

            tk.Label(self.frame, text=f"שנים להחזר (תרחיש {i+1}):").grid(row=row_idx, column=0, sticky="e")
            years_entry = tk.Entry(self.frame)
            years_entry.grid(row=row_idx, column=1)
            self.years_entries.append(years_entry)
            row_idx += 1

        self.calc_button = tk.Button(self.frame, text="חשב", command=self.calculate)
        self.calc_button.grid(row=row_idx, column=0, columnspan=2, pady=10)
        row_idx += 1

        self.export_button = tk.Button(self.frame, text="ייצא לטבלה (CSV)", command=self.export_to_csv)
        self.export_button.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.tax_label = tk.Label(self.frame, text="")
        self.tax_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.downpayment_label = tk.Label(self.frame, text="")
        self.downpayment_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.lawyer_fee_label = tk.Label(self.frame, text="")
        self.lawyer_fee_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.broker_fee_label = tk.Label(self.frame, text="")
        self.broker_fee_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.total_funds_label = tk.Label(self.frame, text="", font=("Arial", 12, "bold"))
        self.total_funds_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.price_per_meter_label = tk.Label(self.frame, text="")
        self.price_per_meter_label.grid(row=row_idx, column=0, columnspan=2)
        row_idx += 1

        self.table = ttk.Treeview(self.frame, columns=("loan", "rate", "years", "monthly", "interest", "total"), show="headings", height=6)
        self.table.grid(row=row_idx, column=0, columnspan=2, pady=15, sticky='nsew')

        self.table.heading("loan", text="סכום הלוואה (₪)")
        self.table.heading("rate", text="ריבית שנתית (%)")
        self.table.heading("years", text="שנים להחזר")
        self.table.heading("monthly", text="תשלום חודשי (₪)")
        self.table.heading("interest", text="סה\"כ ריבית (₪)")
        self.table.heading("total", text="סה\"כ תשלום כולל (₪)")

        self.table.column("loan", width=120, anchor="center")
        self.table.column("rate", width=100, anchor="center")
        self.table.column("years", width=100, anchor="center")
        self.table.column("monthly", width=130, anchor="center")
        self.table.column("interest", width=130, anchor="center")
        self.table.column("total", width=130, anchor="center")

        row_idx += 1

        self.figure_list = []
        self.ax_list = []
        self.canvas_list = []
        for i in range(3):
            fig = plt.Figure(figsize=(5.5, 2.5), dpi=100)
            ax = fig.add_subplot(111)
            canvas = FigureCanvasTkAgg(fig, self.frame)
            canvas.get_tk_widget().grid(row=row_idx + i, column=0, columnspan=2, pady=5)
            self.figure_list.append(fig)
            self.ax_list.append(ax)
            self.canvas_list.append(canvas)

        self.df_list = [None, None, None]

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def calculate(self):
        try:
            price = float(self.price_entry.get())
            area_text = self.area_entry.get()
            area = float(area_text) if area_text else None
            ltv = float(self.ltv_entry.get())

            rates = []
            years = []
            for i in range(3):
                rate_val = self.rate_entries[i].get()
                years_val = self.years_entries[i].get()
                if rate_val and years_val:
                    rates.append(float(rate_val))
                    years.append(int(years_val))
                else:
                    rates.append(None)
                    years.append(None)
        except ValueError:
            return

        purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
        lawyer_fee = estimate_lawyer_fee(price)
        broker_fee = 0 if self.skip_broker_var.get() else estimate_broker_fee(price)

        loan_amount = price * (ltv / 100)
        down_payment = price - loan_amount

        total_needed = down_payment + purchase_tax + lawyer_fee + broker_fee

        self.tax_label.config(text=f"מס רכישה משוער: {purchase_tax:,.0f} ₪")
        self.downpayment_label.config(text=f"הון עצמי נדרש: {down_payment:,.0f} ₪")
        self.lawyer_fee_label.config(text=f"עלות עורך דין משוערת: {lawyer_fee:,.0f} ₪")
        self.broker_fee_label.config(text=f"עלות מתווך משוערת: {broker_fee:,.0f} ₪")
        self.total_funds_label.config(text=f"סה\"כ הון דרוש (כולל הון עצמי, עו\"ד, מתווך ומס רכישה): {total_needed:,.0f} ₪")

        if area and area > 0:
            price_per_meter = price / area
            self.price_per_meter_label.config(text=f"מחיר למטר מרובע: {price_per_meter:,.2f} ₪")
        else:
            self.price_per_meter_label.config(text="")

        # ניקוי הטבלה לפני הכנסת שורות חדשות
        for i in self.table.get_children():
            self.table.delete(i)

        for i in range(3):
            if rates[i] is not None and years[i] is not None:
                df = generate_amortization_df(loan_amount, rates[i], years[i])
                self.df_list[i] = df
                total_interest = df["ריבית"].sum()
                total_payment = df["תשלום חודשי"].sum()
                monthly_payment = calculate_monthly_payment(loan_amount, rates[i], years[i])

                self.table.insert("", "end", values=(
                    f"{loan_amount:,.0f}",
                    f"{rates[i]:.2f}",
                    f"{years[i]}",
                    f"{monthly_payment:,.0f}",
                    f"{total_interest:,.0f}",
                    f"{total_payment:,.0f}",
                ))

                ax = self.ax_list[i]
                ax.clear()
                ax.plot(df["חודש"], df["קרן"], label="קרן", color="green")
                ax.plot(df["חודש"], df["ריבית"], label="ריבית", color="red")
                ax.invert_yaxis()
                ax.set_title(f"גרף תשלום חודשי - תרחיש {i+1}")
                ax.set_xlabel("חודש")
                ax.set_ylabel("₪")
                ax.legend()
                ax.grid(True)
                ax.set_xlim(left=1)
                self.canvas_list[i].draw()
            else:
                self.df_list[i] = None
                self.table.insert("", "end", values=("אין נתונים",) * 6)
                self.ax_list[i].clear()
                self.canvas_list[i].draw()

    def export_to_csv(self):
        if any(self.df_list):
            file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if file_path:
                with open(file_path, "w", encoding="utf-8-sig") as f:
                    f.write("מס רכישה משוער,הון עצמי נדרש,עלות עורך דין,עלות מתווך,סה\"כ הון דרוש,מחיר למטר מרובע\n")
                    price = float(self.price_entry.get())
                    purchase_tax = 0 if self.skip_tax_var.get() else calculate_purchase_tax(price)
                    loan_amount = price * (float(self.ltv_entry.get()) / 100)
                    down_payment = price - loan_amount
                    lawyer_fee = estimate_lawyer_fee(price)
                    broker_fee = 0 if self.skip_broker_var.get() else estimate_broker_fee(price)
                    total_needed = down_payment + purchase_tax + lawyer_fee + broker_fee
                    area_text = self.area_entry.get()
                    area = float(area_text) if area_text else None
                    price_per_meter = (price / area) if area and area > 0 else ""
                    f.write(f"{purchase_tax:.2f},{down_payment:.2f},{lawyer_fee:.2f},{broker_fee:.2f},{total_needed:.2f},{price_per_meter}\n\n")

                    for i, df in enumerate(self.df_list):
                        if df is not None:
                            f.write(f"תרחיש {i+1}\n")
                            df.to_csv(f, index=False, encoding="utf-8-sig")
                            f.write("\n\n")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("950x850")
    app = MortgageApp(root)
    root.mainloop()
