""" Installation requirements: pandas, openpyxl """

import tkinter as tk
from tkinter import ttk
import pandas as pd
import os

def create_subscript_label(master, x, y, text_elements, main_font):
    """Create a subscript label using a masked Text widget."""
    lbl = tk.Text(master, height=1, width=70, borderwidth=0, bg="white", font=main_font, highlightthickness=0, pady=2)
    lbl.place(x=x, y=y-2) 
    
    sub_font = ("Arial", 7, "bold") 
    lbl.tag_configure("sub", offset="-2", font=sub_font)
    
    for text, tag in text_elements:
        if tag == "sub":
            lbl.insert(tk.END, text, "sub")
        else:
            lbl.insert(tk.END, text)
            
    lbl.config(state="disabled")

def extract_values(raw_data, target_string, month_ita):
    """Helper function to extract ventilated and non-ventilated data."""
    col_q = raw_data.iloc[:, 16].fillna("").astype(str)
    try:
        row_q = col_q[col_q == target_string].index[0]
    except IndexError:
        raise ValueError("Header not found in Col Q.")
    
    range_q = raw_data.iloc[row_q + 2 : min(row_q + 12, len(raw_data)), 16].fillna("").astype(str)
    try:
        idx_month_q = range_q[range_q == month_ita].index[0]
    except IndexError:
        raise ValueError(f"Month {month_ita} not found in Col Q range.")
        
    val_vent = float(raw_data.iloc[idx_month_q, 18])

    col_v = raw_data.iloc[:, 21].fillna("").astype(str)
    try:
        row_v = col_v[col_v == target_string].index[0]
    except IndexError:
        raise ValueError("Header not found in Col V.")
        
    range_v = raw_data.iloc[row_v + 2 : min(row_v + 12, len(raw_data)), 21].fillna("").astype(str)
    try:
        idx_month_v = range_v[range_v == month_ita].index[0]
    except IndexError:
        raise ValueError(f"Month {month_ita} not found in Col V range.")
        
    val_non_vent = float(raw_data.iloc[idx_month_v, 23])

    return val_vent, val_non_vent

def run_calculation():
    """Main callback to calculate cooling energy."""
    txt_results.place(x=40, y=410, width=800, height=170)
    txt_results.config(state="normal")
    txt_results.delete("1.0", tk.END)
    txt_results.insert(tk.END, "Calculation in progress. Interpolating data...\n", "bold")
    txt_results.config(state="disabled")
    root.update()

    try:
        val_Wc = float(dd_Wc.get())
        val_Hw = float(dd_Hw.get())
        val_eps = dd_eps.get()
        val_U = float(dd_U.get())
        val_Ori = dd_Ori.get()
        val_City = dd_City.get()
        val_Month = dd_Month.get()

        orientation_map = {'East': 'Est', 'West': 'Ovest'}
        month_map = {
            'May': 'Maggio', 'June': 'Giugno', 'July': 'Luglio', 
            'August': 'Agosto', 'September': 'Settembre'
        }
        val_Ori_ITA = orientation_map[val_Ori]
        val_Month_ITA = month_map[val_Month]

        eps_clean = val_eps.replace(' ', '').replace('.', '').replace('-', '_')

        folder_path = os.path.dirname(os.path.abspath(__file__))
        sheet_name = f"{val_Wc:g}, {val_Hw:g}, {val_Ori_ITA.lower()}"
        target_string = f"Cooling load U = {val_U:g} W/(m2*K)"

        city_data = cdd_data[cdd_data['City'] == val_City]
        if city_data.empty:
            raise ValueError(f'City "{val_City}" not found in CDD Database.')
        
        cdd_City = city_data['CDD_26'].values[0]
        cdd_TO = cdd_data.loc[cdd_data['City'] == 'Torino', 'CDD_26'].values[0]
        cdd_ME = cdd_data.loc[cdd_data['City'] == 'Messina', 'CDD_26'].values[0]

        file_TO = os.path.join(folder_path, f'Torino_{eps_clean}.xlsx')
        data_TO = pd.read_excel(file_TO, sheet_name=sheet_name, header=None)
        vent_TO, nonVent_TO = extract_values(data_TO, target_string, val_Month_ITA)

        file_ME = os.path.join(folder_path, f'Messina_{eps_clean}.xlsx')
        data_ME = pd.read_excel(file_ME, sheet_name=sheet_name, header=None)
        vent_ME, nonVent_ME = extract_values(data_ME, target_string, val_Month_ITA)

        ratio = (cdd_City - cdd_TO) / (cdd_ME - cdd_TO)
        
        vent_City = vent_TO + (vent_ME - vent_TO) * ratio
        nonVent_City = nonVent_TO + (nonVent_ME - nonVent_TO) * ratio

        vent_City = max(0, vent_City)
        nonVent_City = max(0, nonVent_City)

        txt_results.config(state="normal")
        txt_results.delete("1.0", tk.END)
        
        txt_results.insert(tk.END, "Input Parameters\n", "bold")
        txt_results.insert(tk.END, f"City: {val_City}\nMonth: {val_Month}\nCavity Depth: {val_Wc:g} cm\nBuilding height: {val_Hw:g} m\nFaçade's Solar absorptance and infrared emissivity α_sol - ε_inf: {val_eps}\nExposure: {val_Ori}\nThermal transmittance: {val_U:g} W/(m² K)\n\n")
        
        txt_results.insert(tk.END, "Cooling Energy Results\n", "bold")
        txt_results.insert(tk.END, f"Ventilated Case: ")
        txt_results.insert(tk.END, f"{vent_City:.6f} kWh/(m² day)\n", "bold")
        txt_results.insert(tk.END, f"Non-Ventilated Case: ")
        txt_results.insert(tk.END, f"{nonVent_City:.6f} kWh/(m² day)\n", "bold")
        
        txt_results.config(state="disabled")

    except Exception as e:
        txt_results.config(state="normal")
        txt_results.delete("1.0", tk.END)
        txt_results.insert(tk.END, "Error:\n", "error")
        txt_results.insert(tk.END, str(e))
        txt_results.config(state="disabled")

class AnimatedGIF(tk.Label):
    """Class to manage GIF animation """
    def __init__(self, master, path):
        super().__init__(master, bg="white")
        self.frames = []
        try:
            i = 0
            while True:
                frame = tk.PhotoImage(file=path, format=f"gif -index {i}")
                self.frames.append(frame.subsample(2, 2))
                i += 1
        except tk.TclError:
            pass 
        
        if self.frames:
            self.config(image=self.frames[0])
            self.delay = 10000 
            self.idx = 0
            self.animate()
        else:
            self.config(text="GIF not found")

    def animate(self):
        self.idx = (self.idx + 1) % len(self.frames)
        self.config(image=self.frames[self.idx])
        self.after(self.delay, self.animate)

# App Initialization
if __name__ == "__main__":
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base_dir = os.getcwd() 
        
    db_path = os.path.join(base_dir, 'CDD_Database.xlsx')

    try:
        cdd_data = pd.read_excel(db_path)
        city_list = cdd_data['City'].tolist()
    except Exception as e:
        city_list = ["Database Not Found"]
        print(f"Failed to load CDD_Database.xlsx: {e}")

    root = tk.Tk()
    root.title("Ventilated Façade Energy Calculator")
    root.geometry("1000x670") 
    root.configure(bg="white") 

    font_bold = ("Arial", 10, "bold")
    font_normal = ("Arial", 10)
    
    root.option_add('*TCombobox*Listbox.font', font_normal)

    tk.Label(root, text="Select input parameters: ", font=("Arial", 11, "bold"), fg="#2C3E50", bg="white").place(x=40, y=15)

    create_subscript_label(root, 40, 60, [
        ("Cavity Depth D", "normal"), 
        ("c", "sub"), 
        (" [cm]:", "normal")
    ], font_bold)
    dd_Wc = ttk.Combobox(root, values=['4', '8'], state="readonly", font=font_normal)
    dd_Wc.current(0)
    dd_Wc.place(x=470, y=60, width=140)

    create_subscript_label(root, 40, 100, [
        ("Wall Height H", "normal"), 
        ("w", "sub"), 
        (" [m]:", "normal")
    ], font_bold)
    dd_Hw = ttk.Combobox(root, values=['6.6', '10'], state="readonly", font=font_normal)
    dd_Hw.current(0)
    dd_Hw.place(x=470, y=100, width=140)

    create_subscript_label(root, 40, 140, [
        ("Façade's Solar absorptance and infrared emissivity α", "normal"), 
        ("sol", "sub"), 
        (" - ε", "normal"),
        ("inf", "sub"),
        (":", "normal")
    ], font_bold)
    dd_eps = ttk.Combobox(root, values=['0.5 - 0.5', '0.5 - 0.9', '0.9 - 0.9'], state="readonly", font=font_normal)
    dd_eps.current(0)
    dd_eps.place(x=470, y=140, width=140)

    tk.Label(root, text="Thermal Transmittance U [W/(m²K)]:", font=font_bold, bg="white").place(x=40, y=180)
    dd_U = ttk.Combobox(root, values=['0.2', '0.5', '0.8'], state="readonly", font=font_normal)
    dd_U.current(0)
    dd_U.place(x=470, y=180, width=140)

    tk.Label(root, text="Orientation:", font=font_bold, bg="white").place(x=40, y=220)
    dd_Ori = ttk.Combobox(root, values=['East', 'West'], state="readonly", font=font_normal)
    dd_Ori.current(0)
    dd_Ori.place(x=470, y=220, width=140)

    tk.Label(root, text="City:", font=font_bold, bg="white").place(x=40, y=260)
    dd_City = ttk.Combobox(root, values=city_list, font=font_normal) 
    if city_list and city_list[0] != "Database Not Found":
        dd_City.current(0)
    dd_City.place(x=470, y=260, width=140)

    tk.Label(root, text="Month:", font=font_bold, bg="white").place(x=40, y=300)
    dd_Month = ttk.Combobox(root, values=['May', 'June', 'July', 'August', 'September'], state="readonly", font=font_normal)
    dd_Month.current(0)
    dd_Month.place(x=470, y=300, width=140)

    # GIF
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base_dir = os.getcwd() 
        
    gif_path = os.path.join(base_dir, 'gif_facade.gif')
    lbl_gif = AnimatedGIF(root, gif_path)
    lbl_gif.place(x=640, y=30) 

    # BUTTON STYLE
    colore_azzurro = "#4A90E2" 
    colore_hover = "#357ABD" 

    def on_enter(e):
        btn_calc['background'] = colore_hover

    def on_leave(e):
        btn_calc['background'] = colore_azzurro

    btn_calc = tk.Button(
        root, 
        text="Calculate Cooling Energy", 
        font=("Arial", 11, "bold"), 
        bg=colore_azzurro, 
        fg="white", 
        activebackground=colore_hover, 
        activeforeground="white", 
        relief="flat", 
        borderwidth=0, 
        cursor="hand2", 
        command=run_calculation
    )
    btn_calc.place(x=250, y=360, width=200, height=40)
    
    btn_calc.bind("<Enter>", on_enter)
    btn_calc.bind("<Leave>", on_leave)

    # Results Area 
    txt_results = tk.Text(root, font=font_normal, bg="white", wrap=tk.WORD)
    txt_results.tag_configure("bold", font=("Arial", 10, "bold"))
    txt_results.tag_configure("italic", font=("Arial", 10, "italic"))
    txt_results.tag_configure("error", font=("Arial", 10, "bold"), foreground="red")
    txt_results.config(state="disabled") 

    root.mainloop()