
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
main_tkinter_reconciliation_app.py

Tkinter launcher for the Trade Reconciliation project.

Three browsed inputs:
1. Internal Trades CSV
2. Market Prices CSV (optional)
3. Existing Reconciliation Detail CSV (optional)

Behavior:
- If reconciliation detail is provided, only Tableau outputs are built.
- Else if market prices are provided, reconciliation + Tableau outputs are run.
- Else market prices are generated from Alpha Vantage, then reconciliation + Tableau outputs are run.

Assumptions:
- This file lives in the same folder as:
    download_market_input.py
    run_reconciliation.py
    build_tableau_outputs.py
- Outputs are written under ./output
"""

from __future__ import annotations

import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


APP_TITLE = "Trade Reconciliation Automation"
WINDOW_SIZE = "1180x760"


class TkinterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.configure(bg="#F7F9FC")
        self.resizable(False, False)

        self.base_dir = Path(__file__).resolve().parent
        self.output_dir = self.base_dir / "output"

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._build_ui()
        self.after(150, self._drain_log_queue)

    def _build_ui(self) -> None:
        title = tk.Label(
            self,
            text="Trade Reconciliation Automation",
            font=("Helvetica", 22, "bold"),
            bg="#F7F9FC",
            fg="#102A43",
        )
        title.pack(pady=(18, 6))

        subtitle = tk.Label(
            self,
            text=(
                "Browse the project files, click Run Automation, and generate "
                "Tableau-ready CSV outputs."
            ),
            font=("Helvetica", 11),
            bg="#F7F9FC",
            fg="#486581",
        )
        subtitle.pack(pady=(0, 14))

        form = tk.Frame(self, bg="#FFFFFF", bd=1, relief="solid")
        form.pack(fill="x", padx=22, pady=(0, 12))

        self.internal_file_var = tk.StringVar()
        self.market_file_var = tk.StringVar()
        self.recon_file_var = tk.StringVar()
        self.api_key_var = tk.StringVar(value="2B1V00IN53XV5Z0A")
        self.abs_tol_var = tk.StringVar(value="1.0")
        self.pct_tol_var = tk.StringVar(value="1.0")

        row = 0
        self._add_file_row(
            parent=form,
            row=row,
            label="1. Internal Trades CSV *",
            variable=self.internal_file_var,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )

        row += 1
        self._add_file_row(
            parent=form,
            row=row,
            label="2. Market Prices CSV (optional)",
            variable=self.market_file_var,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )

        row += 1
        self._add_file_row(
            parent=form,
            row=row,
            label="3. Reconciliation Detail CSV (optional)",
            variable=self.recon_file_var,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )

        row += 1
        self._add_entry_row(
            parent=form,
            row=row,
            label="Alpha Vantage API Key",
            variable=self.api_key_var,
            show="",
        )

        row += 1
        tol_frame = tk.Frame(form, bg="#FFFFFF")
        tol_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=18, pady=10)
        tol_frame.columnconfigure(1, weight=1)
        tol_frame.columnconfigure(3, weight=1)

        tk.Label(
            tol_frame,
            text="Absolute Tolerance",
            font=("Helvetica", 10, "bold"),
            bg="#FFFFFF",
            fg="#243B53",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        tk.Entry(
            tol_frame,
            textvariable=self.abs_tol_var,
            font=("Helvetica", 10),
            width=12,
        ).grid(row=0, column=1, sticky="w")

        tk.Label(
            tol_frame,
            text="Percentage Tolerance",
            font=("Helvetica", 10, "bold"),
            bg="#FFFFFF",
            fg="#243B53",
        ).grid(row=0, column=2, sticky="w", padx=(28, 8))

        tk.Entry(
            tol_frame,
            textvariable=self.pct_tol_var,
            font=("Helvetica", 10),
            width=12,
        ).grid(row=0, column=3, sticky="w")

        buttons = tk.Frame(self, bg="#F7F9FC")
        buttons.pack(fill="x", padx=22, pady=(0, 10))

        self.run_button = tk.Button(
            buttons,
            text="Run Automation",
            font=("Helvetica", 11, "bold"),
            bg="#1F6FEB",
            fg="white",
            activebackground="#1558C0",
            activeforeground="white",
            width=18,
            command=self.start_run,
        )
        self.run_button.pack(side="left", padx=(0, 10))

        tk.Button(
            buttons,
            text="Open Output Folder",
            font=("Helvetica", 11),
            width=18,
            command=self.open_output_folder,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            buttons,
            text="Clear Log",
            font=("Helvetica", 11),
            width=14,
            command=self.clear_log,
        ).pack(side="left")

        info = tk.Label(
            self,
            text=(
                "Run order: download_market_input.py → run_reconciliation.py → "
                "build_tableau_outputs.py"
            ),
            font=("Helvetica", 10),
            bg="#F7F9FC",
            fg="#52606D",
        )
        info.pack(anchor="w", padx=24, pady=(2, 8))

        self.log_text = scrolledtext.ScrolledText(
            self,
            wrap="word",
            height=24,
            width=140,
            font=("Consolas", 10),
            bg="#0B1220",
            fg="#E6EEF8",
            insertbackground="white",
        )
        self.log_text.pack(fill="both", expand=True, padx=22, pady=(0, 18))
        self.log_text.config(state="disabled")

    def _add_file_row(
        self,
        parent: tk.Widget,
        row: int,
        label: str,
        variable: tk.StringVar,
        filetypes: list[tuple[str, str]],
    ) -> None:
        frame = tk.Frame(parent, bg="#FFFFFF")
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=18, pady=10)
        frame.columnconfigure(1, weight=1)

        tk.Label(
            frame,
            text=label,
            font=("Helvetica", 10, "bold"),
            bg="#FFFFFF",
            fg="#243B53",
            width=30,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        tk.Entry(
            frame,
            textvariable=variable,
            font=("Helvetica", 10),
            width=82,
        ).grid(row=0, column=1, sticky="ew")

        tk.Button(
            frame,
            text="Browse",
            width=12,
            command=lambda: self.browse_file(variable, filetypes),
        ).grid(row=0, column=2, padx=(12, 0))

    def _add_entry_row(
        self,
        parent: tk.Widget,
        row: int,
        label: str,
        variable: tk.StringVar,
        show: str = "",
    ) -> None:
        frame = tk.Frame(parent, bg="#FFFFFF")
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=18, pady=10)
        frame.columnconfigure(1, weight=1)

        tk.Label(
            frame,
            text=label,
            font=("Helvetica", 10, "bold"),
            bg="#FFFFFF",
            fg="#243B53",
            width=30,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        tk.Entry(
            frame,
            textvariable=variable,
            font=("Helvetica", 10),
            width=82,
            show=show,
        ).grid(row=0, column=1, sticky="ew")

    def browse_file(
        self,
        variable: tk.StringVar,
        filetypes: list[tuple[str, str]],
    ) -> None:
        selected = filedialog.askopenfilename(
            title="Select file",
            filetypes=filetypes,
            initialdir=str(self.base_dir),
        )
        if selected:
            variable.set(selected)

    def log(self, message: str) -> None:
        self.log_queue.put(message)

    def _drain_log_queue(self) -> None:
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(150, self._drain_log_queue)

    def clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def start_run(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Already Running", "Automation is already in progress.")
            return

        internal_file = self.internal_file_var.get().strip()
        if not internal_file:
            messagebox.showerror("Missing Input", "Please select Internal Trades CSV.")
            return

        self.run_button.config(state="disabled")
        self.clear_log()
        self.worker_thread = threading.Thread(target=self.run_pipeline, daemon=True)
        self.worker_thread.start()

    def run_pipeline(self) -> None:
        try:
            internal_file = Path(self.internal_file_var.get().strip())
            market_file = Path(self.market_file_var.get().strip()) if self.market_file_var.get().strip() else None
            recon_file = Path(self.recon_file_var.get().strip()) if self.recon_file_var.get().strip() else None
            api_key = self.api_key_var.get().strip()
            abs_tol = self.abs_tol_var.get().strip() or "1.0"
            pct_tol = self.pct_tol_var.get().strip() or "1.0"

            project_dir = self.base_dir
            output_dir = project_dir / "output"
            output_dir.mkdir(parents=True, exist_ok=True)

            download_script = project_dir / "download_market_input.py"
            recon_script = project_dir / "run_reconciliation.py"
            tableau_script = project_dir / "build_tableau_outputs.py"

            for script_path in [download_script, recon_script, tableau_script]:
                if not script_path.exists():
                    raise FileNotFoundError(f"Required script not found: {script_path}")

            if not internal_file.exists():
                raise FileNotFoundError(f"Internal file not found: {internal_file}")

            self.log(">>> Project execution started")
            self.log(f">>> Project folder: {project_dir}")
            self.log(f">>> Output folder: {output_dir}")

            final_market_file = output_dir / "market_prices.csv"
            final_recon_file = output_dir / "reconciliation_detail.csv"
            final_exceptions_file = output_dir / "exceptions_detail.csv"
            final_metadata_file = output_dir / "run_metadata.json"

            # Case 1: user already has reconciliation_detail.csv
            if recon_file:
                if not recon_file.exists():
                    raise FileNotFoundError(f"Reconciliation detail file not found: {recon_file}")

                self.log(">>> Existing reconciliation_detail.csv provided; skipping market download and reconciliation.")
                shutil.copy2(recon_file, final_recon_file)
                self.log(f"[INFO] Copied reconciliation detail to {final_recon_file}")

                self._run_subprocess(
                    [
                        sys.executable,
                        str(tableau_script),
                    ],
                    cwd=project_dir,
                    step_name="Build Tableau Outputs",
                )

            else:
                # Case 2: use supplied market prices or generate them
                if market_file:
                    if not market_file.exists():
                        raise FileNotFoundError(f"Market prices file not found: {market_file}")

                    self.log(">>> Existing market_prices.csv provided; skipping market download.")
                    shutil.copy2(market_file, final_market_file)
                    self.log(f"[INFO] Copied market prices to {final_market_file}")

                else:
                    if not api_key:
                        raise ValueError("API key is required when Market Prices CSV is not provided.")

                    self._run_subprocess(
                        [
                            sys.executable,
                            str(download_script),
                            "--input-file",
                            str(internal_file),
                            "--output-file",
                            str(final_market_file),
                            "--api-key",
                            api_key,
                        ],
                        cwd=project_dir,
                        step_name="Download Market Input",
                    )

                # Run reconciliation
                self._run_subprocess(
                    [
                        sys.executable,
                        str(recon_script),
                        "--internal-file",
                        str(internal_file),
                        "--market-file",
                        str(final_market_file),
                        "--output-file",
                        str(final_recon_file),
                        "--exceptions-file",
                        str(final_exceptions_file),
                        "--metadata-file",
                        str(final_metadata_file),
                        "--abs-tolerance",
                        abs_tol,
                        "--pct-tolerance",
                        pct_tol,
                    ],
                    cwd=project_dir,
                    step_name="Run Reconciliation",
                )

                # Build Tableau outputs
                self._run_subprocess(
                    [
                        sys.executable,
                        str(tableau_script),
                    ],
                    cwd=project_dir,
                    step_name="Build Tableau Outputs",
                )

            self.log("")
            self.log(">>> Project execution completed successfully")
            self.log(">>> Tableau-ready output files generated under ./output")
            self.log("    - market_prices.csv")
            self.log("    - reconciliation_detail.csv")
            self.log("    - exceptions_detail.csv")
            self.log("    - run_metadata.json")
            self.log("    - kpi_summary.csv")
            self.log("    - status_summary.csv")
            self.log("    - severity_summary.csv")
            self.log("    - symbol_breaks.csv")

            messagebox.showinfo(
                "Success",
                "Automation completed successfully.\n\nOutput files are ready for Tableau.",
            )

        except Exception as exc:
            self.log(f"[ERROR] {exc}")
            messagebox.showerror("Execution Failed", str(exc))

        finally:
            self.run_button.config(state="normal")

    def _run_subprocess(self, cmd: list[str], cwd: Path, step_name: str) -> None:
        self.log("")
        self.log(f">>> Step started: {step_name}")
        self.log(f"[CMD] {' '.join(cmd)}")

        process = subprocess.Popen(
            cmd,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )

        assert process.stdout is not None
        for line in process.stdout:
            self.log(line.rstrip())

        return_code = process.wait()
        if return_code != 0:
            raise RuntimeError(f"{step_name} failed with return code {return_code}")

        self.log(f">>> Step completed: {step_name}")

    def open_output_folder(self) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)

        system_name = platform.system()
        if system_name == "Windows":
            os.startfile(self.output_dir)  # type: ignore[attr-defined]
        elif system_name == "Darwin":
            subprocess.run(["open", str(self.output_dir)], check=False)
        else:
            subprocess.run(["xdg-open", str(self.output_dir)], check=False)


if __name__ == "__main__":
    app = TkinterApp()
    app.mainloop()


















#!/usr/bin/env python3
# # -*- coding: utf-8 -*-
# """
# Created on Fri Apr 10 22:57:56 2026

# @author: komalwavhal
# """


# #try:
# #    import docx
# #except:
# #    os.system('python -m pip install -i python-docx')
# #    import docx
   
 
# import os
# import tkinter as tk
# from tkinter import ttk
# from PIL import Image,ImageTk
# import threading
# import time
# from tkinter import scrolledtext
# from tkinter import messagebox
# LARGEFONT =("Verdana", 35) 
# import requests
# from io import BytesIO
# import webbrowser
# from urllib.parse import quote_plus
# import pandas as pd 

# class tkinterApp(tk.Tk):
    
#     def get_base_path(self):
#         from pathlib import Path
#         return Path(__file__).resolve().parent
        
#     # __init__ function for class tkinterApp   
#     def __init__(self, *args, **kwargs):
#         tk.Tk.__init__(self, *args, **kwargs)

#         container = tk.Frame(self)
#         container.pack(side="top", fill="both", expand=True)

#         container.grid_rowconfigure(0, weight=1)
#         container.grid_columnconfigure(0, weight=1)

#         self.frames = {}

#         for F in (StartPage, Page1):
#             frame = F(container, self)
#             self.frames[F] = frame
#             frame.grid(row=0, column=0, sticky="nsew")

#         self.show_frame(StartPage)
       
    
    
#     # to display the current frame passed as  parameter
#     def show_frame(self, cont): 
    
#         frame = self.frames[cont]
#         frame.tkraise() 
            
#         BASE_DIR = self.get_base_path()
#         print(' system BASE_DIR -  ' , BASE_DIR)
        
#         try:
#             ### for Windows users
#             desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
#             print(desktop)
#             Parent_Folder_Path = desktop + "/AAI_551_Project_Repository"
#             print(Parent_Folder_Path)
        
#         except: 
#             ### for Mac users
#             desktop = os.path.expanduser("~/Desktop")
#             desktop = os.path.normpath(os.path.expanduser("~/Desktop"))
#             Parent_Folder_Path = desktop + "/AAI_551_Project_Repository"
#             print(Parent_Folder_Path)
            
#         Parent_Folder_Path = Parent_Folder_Path
        

# # first window frame startpage 
# class StartPage(tk.Frame):

#     def display_text(self, text,tag):

#         self.text.tag_config('Error', background="white", foreground="red",font=('calibri',12))
#         self.text.tag_config('Normal', background="white",foreground="#819ec4",font=('calibri',12))
        
#         if (tag == 'Error') : 
#             self.text.insert(INSERT, text,'Error')
        			
#         if (tag == 'Normal') : 
#             self.text.insert(INSERT, text,'Normal')
        			 
#         self.text.see("end")  


#     def __init__(self, parent, controller): 
            

#         tk.Frame.__init__(self, parent)

#         BASE_DIR = controller.get_base_path()
#         print("system BASE_DIR -", BASE_DIR)

#         Img_Path = BASE_DIR / "Images"
#         imgFilePath = Img_Path / "pic-11.png"
#         AFFIRM_ICON_imgFilePath = Img_Path / "AFFIRM_ICON.png"

#         img = Image.open(imgFilePath)
#         photo = ImageTk.PhotoImage(img)

#         lab1 = ttk.Label(self, image=photo)
#         lab1.place(x=0)
#         lab1.image = photo

#         affirm_Button_image = Image.open(AFFIRM_ICON_imgFilePath)
#         photo2 = ImageTk.PhotoImage(affirm_Button_image)

#         affirm_Button = tk.Button(
#             self,
#             image=photo2,
#             command=lambda: controller.show_frame(Page1)
#         )
#         affirm_Button.place(x=667, y=420)
#         affirm_Button.image = photo2

#         self.configure(background='#FFFFFF')
        
        
#         ######### image path   ###########
#         from pathlib import Path
        
#         BASE_DIR = Path(__file__).resolve().parent
#         Img_Path = BASE_DIR / "Images"
        
#         userManual_Button_imgFilePath = Img_Path / "UserManual.png"
#         imgFilePath = Img_Path / "pic-11.png"
#         AFFIRM_ICON_imgFilePath = Img_Path / "AFFIRM_ICON.png"
        
#         print("BASE_DIR:", BASE_DIR)
#         print("Img_Path:", Img_Path)
#         print("Image exists:", imgFilePath.exists())
        


# # second window frame page1
# class Page1(tk.Frame):
    
      
#     def domain_changed_1(self,event):
#         self.mainGetData_1 = list()
#         self.mainGetData_1.append(self.domainselected_1.get())
    		 	
        
#     def browsefunc_1(self):
#         from pathlib import Path
#         import os
#         import platform
#         import subprocess
    
#         BASE_DIR = Path(__file__).resolve().parent
#         User_Manual_Path = BASE_DIR / "User_Manual"
    
#         print("BASE_DIR:", BASE_DIR)
#         print("User_Manual_Path:", User_Manual_Path)
#         print("Exists:", User_Manual_Path.exists())
    
#         if not User_Manual_Path.exists():
#             messagebox.showerror("Error", f"Folder not found:\n{User_Manual_Path}")
#             return
    
#         system_name = platform.system()
    
#         if system_name == "Windows":
#             os.startfile(User_Manual_Path)
#         elif system_name == "Darwin":  # macOS
#             subprocess.run(["open", str(User_Manual_Path)])
#         else:  # Linux
#             subprocess.run(["xdg-open", str(User_Manual_Path)])
            
            
#     def browsefunc_1(self): 
         
#         win_user = 0
#         import os
#         try:
#             ### for Windows users
#             desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
#             print(desktop)
#             Parent_Folder_Path = desktop + "/AAI_551_Project_Repository"
#             print(Parent_Folder_Path)
        
#         except: 
#             win_user = 1            
#             ### for Mac users
#             desktop = os.path.expanduser("~/Desktop")
#             desktop = os.path.normpath(os.path.expanduser("~/Desktop"))
#             Parent_Folder_Path = desktop + "/AAI_551_Project_Repository"
#             print(Parent_Folder_Path)
            
          
#     #     Parent_Folder_Path = Parent_Folder_Path
#     #     User_Manual_Path = Parent_Folder_Path +'\\' + 'User_Manual' 
         
#     #     export_File_path =  User_Manual_Path
#     #     path = export_File_path
#     #     path = os.path.realpath(path)
#     #     os.startfile(path)
        
         
#     def get_cities_by_country(self, country_name, file_path):
#         # Read the Excel file into a DataFrame
#         df = pd.read_excel(file_path, engine='openpyxl')
          
#         # Filter the DataFrame based on the country name
#         filtered_df = df[df['country'] == country_name]
#         return filtered_df[['city', 'lat', 'lng']]
    
    
    
#     # Step 3: Update the cities dropdown when a country is selected
#     def update_city_dropdown(self,event,   country_name, file_path, city_combobox, lat_entry, lng_entry):
        
#         # Get the cities for the selected country
#         cities_df = self.get_cities_by_country(country_name, file_path)
        
#         # Update the city dropdown options
#         city_combobox['values'] = cities_df['city'].tolist()
#         city_combobox.set('')  # Clear current selection
        
#         # Clear latitude and longitude
#         lat_entry.delete(0, tk.END)
#         lng_entry.delete(0, tk.END)
    
#         # Re-enable the latitude and longitude fields in case they were previously disabled
#         lat_entry.config(state='normal')
#         lng_entry.config(state='normal')
    
    
    
#     # Step 4: Update latitude and longitude when a city is selected
#     def update_lat_lng(self,event, file_path, city_name, lat_entry, lng_entry):
        
#         # Read the Excel file into a DataFrame
#         df = pd.read_excel(file_path, engine='openpyxl')
        
#         # Get the latitude and longitude of the selected city
#         city_info = df[df['city'] == city_name].iloc[0]
#         lat_entry.delete(0, tk.END)
#         lng_entry.delete(0, tk.END)
        
#         # Set latitude and longitude in the entry fields
#         lat_entry.insert(0, city_info['lat'])
#         lng_entry.insert(0, city_info['lng'])
    
#         # Disable the latitude and longitude fields after they are populated
#         lat_entry.config(state='disabled')
#         lng_entry.config(state='disabled')
        
#      ########################################################################################################
#      ########################################################################################################

#     def __init__(self, parent, controller):
#         from pathlib import Path
        
#         BASE_DIR = controller.get_base_path()
#         print("system BASE_DIR -", BASE_DIR)
        
#         Parent_Folder_Path = BASE_DIR
#         Img_Path = BASE_DIR / "Images"
        
#         imgFilePath = Img_Path / "pic-1.png"
#         Bck_Button_imgFilePath = Img_Path / "Bck_Button.png"
#         run_Button_imgFilePath = Img_Path / "run.PNG"
        
#         print("Img_Path:", Img_Path)
#         print("pic-1 exists:", imgFilePath.exists())
#         print("Back button exists:", Bck_Button_imgFilePath.exists())
#         print("Run button exists:", run_Button_imgFilePath.exists())
        
#         ##------------------------------------------------------------------------------------------------------------
          
#         ###-----------------------------------------------------------------------------------------------------------

        		
#         tk.Frame.__init__(self, parent)
        
#         BASE_DIR = controller.get_base_path()
#         print("system BASE_DIR -", BASE_DIR)
        
#         Parent_Folder_Path = BASE_DIR
#         Img_Path = BASE_DIR / "Images"
        
#         imgFilePath = Img_Path / "pic-1.png"
#         Bck_Button_imgFilePath = Img_Path / "Bck_Button.png"
#         run_Button_imgFilePath = Img_Path / "run.PNG"
        
#         print("Img_Path:", Img_Path)
#         print("pic-1 exists:", imgFilePath.exists())
        
#         img = Image.open(imgFilePath)
        
#         photo = ImageTk.PhotoImage(img)
#         lab1 = ttk.Label(self, image=photo)
#         lab1.pack()
#         lab1.place(x=0)
#         lab1.image = photo
        
#         Bck_Button_image = Image.open(Bck_Button_imgFilePath)
#         photo = ImageTk.PhotoImage(Bck_Button_image)
#         Bck_Button = tk.Button(self, image=photo, command=lambda: controller.show_frame(StartPage))
#         Bck_Button.place(x=37, y=30)
#         Bck_Button.image = photo
        
#         self.configure(background='#FFFFFF')
                
#         import getpass
#         import platform
        
#         user = getpass.getuser()
#         machine_name = platform.node()
#         fullname = user
        
#         print("user:", user)
#         print("machine_name:", machine_name)
#         print("fullname:", fullname)
        
#         # fullname = ''

#         # if win_user == 0:	
#         #     print('----------Write down code for windows user name here--------')
#         # else:	            
#         #     import platform
#         #     sys_owner_fname = ( platform.node().replace("s-MacBook-Pro.local",' ') ).lower()
#         #     print('sys_owner_fname' , sys_owner_fname)
#         #     import getpass
#         #     sys_owner_lastname = str( getpass.getuser() )
#         #     sys_owner_lname = sys_owner_lastname.upper()
#         #     replace_val = str(sys_owner_fname.upper()).strip()
#         #     sys_owner_lname = sys_owner_lname.replace( replace_val,'')
#         #     sys_owner_lname = sys_owner_lname.lower()
#         #     print('sys_owner_lname' , sys_owner_lname)

#         #     fullname = sys_owner_fname + '' + sys_owner_lname
#         #     print(fullname)
               
#         #### ----------###(Application Logs)---------------------------------- 
#         self.text = scrolledtext.ScrolledText(self, wrap='word', bg='#FFFFFF', height=3, width=46, font=('calibri',12), relief='ridge')
#         # self.text = scrolledtext.ScrolledText(self,wrap = WORD,bg = '#FFFFFF',  height=3, width=46,  font=('calibri',12),relief= 'ridge') 
#         self.text.place(x=1000,y=700)
        
#         self.configure(background='#FFFFFF') 
          
        
#         ###------------------Select dropdown button ------------------------------------------------- 
        
#         file_path = Parent_Folder_Path / "worldcities.xlsx" 
#         df = pd.read_excel(file_path, engine='openpyxl')
        
#         # Create the country dropdown
#         self.country_names = df['country'].dropna().unique().tolist()
#         self.country_combobox = ttk.Combobox(self, values=self.country_names)
#         # self.country_combobox.set("Select a country")
#         # self.country_combobox.pack(padx=10, pady=120)
#         self.country_combobox.place(x=510,y=200)    
    
#         ###-----------------------------------------------------------------------------------------
             
#         # Create the city dropdown
#         self.city_combobox = ttk.Combobox(self, values=[])
#         # self.city_combobox.set("Select a city")
#         # self.city_combobox.pack(padx=10, pady=70)
#         self.city_combobox.place(x=510,y=230)    
        
#         ###-----------------------------------------------------------------------------------------
             
#         # Create entry fields for latitude and longitude
#         self.lat_label = tk.Label(self, text="Latitude:  ", bg = '#FFFFFF', width=9,  font=('calibri',12),relief= 'ridge')
         
#         self.lat_label.place(x=810,y=206)    
        
#         self.lat_entry = tk.Entry(self)
         
#         self.lat_entry.place(x=910,y=200)    
     
#         ###-----------------------------------------------------------------------------------------
             
#         self.lng_label = tk.Label(self, text="Longitude:", bg = '#FFFFFF', width=9,  font=('calibri',12),relief= 'ridge')
        
#         self.lng_label.place(x=810,y=236)    
        
#         self.lng_entry = tk.Entry(self)
         
#         self.lng_entry.place(x=910,y=235)    
         
#         # Update city dropdown when country is selected
#         self.country_combobox.bind("<<ComboboxSelected>>",  lambda event: self.update_city_dropdown(event, self.country_combobox.get(),  file_path, self.city_combobox, self.lat_entry, self.lng_entry))
    
#         # Update latitude and longitude when city is selected
#         self.city_combobox.bind("<<ComboboxSelected>>",  lambda event: self.update_lat_lng(event, file_path, self.city_combobox.get(), self.lat_entry, self.lng_entry))
 
#         ###-----------------------------------------------------------------------------------------
                 
#         self.Search_label = tk.Label(self, text="Search:     ", bg = '#FFFFFF', width=8,  font=('calibri',12),relief= 'ridge')
#         # lng_label.pack(padx=10, pady=6)
#         self.Search_label.place(x=810,y=270)   
        
#         self.v = tk.IntVar()
#         self.Search_label_entry = tk.Entry(self,text=self.v) 
#         self.Search_label_entry.place(x=910,y=270)    
#         self.v.set(100)  
        
#         ###-----------------------------------------------------------------------------------------
        
         
        
             
#         ###------------------Select radio button -----------------------------------------------------------------------
        

#         ###-----------------------------------------------------------------------------------------
               
        
#         #######Dropdown Button   - for Natural_Attractions_selected
      
#         ###-----------------------------------------------------------------------------------------
               

        
        
        
#         ###-----------------------------------------------------------------------------------------
        
        
        
#         #####RUN Automatoin  
#         runWanderSphere_image = Image.open(run_Button_imgFilePath) 
#         photo = ImageTk.PhotoImage(runWanderSphere_image)        
#         button_Export = tk.Button(self,image=photo,command=  self.Execute_speaker_eng)  ###,bg = "white", bd = 0
#         button_Export.place(x=610,y=310)        
#         button_Export.image = photo  
                  
#         ##### ----------###(THIS STYLE FOR THE PROGRESSBAR)----------------------------------
#         self.style = ttk.Style(self)
        
#         self.style.layout('text.Horizontal.TProgressbar',
#         	 [('Horizontal.Progressbar.trough',
#         	   {'children': [('Horizontal.Progressbar.pbar',
#         					  {'side': 'left', 'sticky': 'ns'})],
#         		'sticky': 'nswe'}),
#         	  ('Horizontal.Progressbar.label', {'sticky': ''})])   
#         self.style.configure('text.Horizontal.TProgressbar', text=' ')
        
        
#         self.configure(background='#FFFFFF')     

#     def Execute_speaker_eng(self) :  
#         # run process in a thread to avoid blocking gui 
#         t = threading.Thread(target=self.execute_main)
#         t.start()
 
    
#     # Function to load and display the image
#     def display_image(self, image_name, image_url, google_map_link, row, col, checkbuttons):
#         try:
#             # Fetch image from the URL
#             response = requests.get(image_url)
#             if response.status_code == 200:
#                 # Open the image using Pillow
#                 img_data = BytesIO(response.content)
#                 img = Image.open(img_data)
                
#                 # Resize image to fit within the window (optional)
#                 img = img.resize((200, 200))  # Adjusted size for 5 images per row
                
#                 # Convert the image for Tkinter compatibility
#                 img_tk = ImageTk.PhotoImage(img)
                
#                 # Create labels to display the Google Map link (above the image)
#                 def open_map():
#                     webbrowser.open(google_map_link)
    
#                 label_map = tk.Label(self.new_window, text="View on Google Maps", fg="blue", bg='#FFFFFF', cursor="hand2", font=("calibri", 15))
#                 label_map.grid(row=row*4, column=col, padx=10, pady=5)
#                 label_map.bind("<Button-1>", lambda e: open_map())  # Bind the click event to open the map link
                
#                 # Create the image label
#                 label_image = tk.Label(self.new_window, image=img_tk)
#                 label_image.image = img_tk  # Keep a reference to avoid garbage collection
#                 label_image.grid(row=row*4+1, column=col, padx=10, pady=10)
                
#                 # Create the image name label (below the image)
#                 label_name = tk.Label(self.new_window, text=image_name, fg="black", bg='#FFFFFF', font=("calibri", 12), wraplength=200)
#                 label_name.grid(row=row*4+2, column=col, padx=10, pady=10)
    
#                 # Create checkbox for the image (below the image name)
#                 var = tk.BooleanVar()
#                 checkbox = tk.Checkbutton(self.new_window, fg="black", bg='#FFFFFF', text=" ", variable=var)
#                 checkbox.grid(row=row*4+3, column=col, padx=10, pady=5)
    
#                 # Store the checkbox state and the place name in the list
#                 checkbuttons.append((var, image_name))  # Store place_name (not link)
    
#                 self.new_window.configure(background='#FFFFFF') 
    
#             else:
#                 print(f"Error: Unable to fetch image {image_name}")
#         except Exception as e:
#             print(f"Failed to load image {image_name}: {str(e)}")
    
#     # Function to generate the Google Map itinerary based on selected checkboxes
#     def generate_itinerary(self, checkbuttons):
#         # Collect the selected locations (place names) based on the checkbox state
#         selected_locations = [place_name for var, place_name in checkbuttons if var.get()]
    
#         if not selected_locations:
#             messagebox.showinfo("No Selection", "Please select at least one place to generate the itinerary.")
#             return
    
#         # Create Google Maps URL with the selected locations as stops
#         base_url = "https://www.google.com/maps/dir/"
#         destination_url = base_url + "/".join([quote_plus(location) for location in selected_locations])
    
#         # Open the generated itinerary in the browser
#         webbrowser.open(destination_url)
     
#     def display_text(self,text,tag):
 

#         self.text.tag_config('Error', background="white", foreground="red",font=('calibri',12))
#         self.text.tag_config('Normal', background="white",foreground="#819ec4",font=('calibri',12))

#         if (tag == 'Error') : 
#             self.text.insert(INSERT, text,'Error')
# 			
#         if (tag == 'Normal') : 
#             self.text.insert(INSERT, text,'Normal')
# 			 
#         self.text.see("end")  
 
#     def open_new_window(self, images_info):
        
#         # Create a new top-level window
#         self.new_window = tk.Toplevel(root)
#         self.new_window.title("Image Display with Google Map Itinerary")
        
#         # Set the size of the new window
#         # self.new_window.geometry("300x200")
         
#         # List to store checkbox variables and their corresponding place names (not links)
#         checkbuttons = []
        
#         # Display all the images with their names, Google Maps links, and checkboxes
#         for index, (image_name, image_url, google_map_link) in enumerate(images_info):
#             row = index // 5  # Determine the row number (5 images per row)
#             col = index % 5   # Determine the column number (5 images per row)
#             self.display_image(image_name, image_url, google_map_link, row, col, checkbuttons)
        
#         # Add a "Generate Google Map Itinerary" button to generate a single itinerary
#         generate_button = tk.Button(self.new_window, text="Generate Google Map Itinerary",fg= 'Black', bg = '#FFFFFF',font=('calibri',15,'bold'), command=lambda: self.generate_itinerary(checkbuttons))
#         generate_button.grid(row=(len(images_info)//5)*4+1, column=0, columnspan=5, pady=20)
        
#         # Start the Tkinter event loop
#         # self.new_window.mainloop()
 
        
#     def execute_main(self):  
       
#         self.text.config(state='normal')
#         self.text.delete("1.0" , "end")
        
#         import datetime
#         now = datetime.datetime.now() 
#         x = str(now.strftime("%Y-%m-%d %H:%M:%S"))  
#         self.display_text('>>> Project execution started ' + '\n\n'  , 'Normal')
#         time.sleep(1)   
        
        
#         city_combobox_Selection = ''
    
#         try:
#             city_combobox_Selection = self.city_combobox.get()
#         except: 
#             city_combobox_Selection = 'Select option'
             
#         print('city_combobox_Selection - ', city_combobox_Selection)
            
           
#         # ####--(Step - 1: Download the flicker dataset)----
#         # Set your Flickr API key and secret
#         api_key = '707492d24c42391a563cddc2bf5e619f'
#         secret_api_key = '52cb54347825468a'
         
        
    
#         try:
#             Natural_Attractions_Selection = self.mainGetData_1
#         except: 
#             Natural_Attractions_Selection = 'Select option'
     
        
#         ########################################################################################
      
#         ##### Importing Custom Libraries for Excel Operations
#         from ExcelUtils import ExcelOperations
#         objExcelOperations = ExcelOperations() 
        
        
#         error_status = ''
      
#         ############################################################################################################################################################################### 
         
          
#         # Join various path components
#         BASE_DIR = self.winfo_toplevel().get_base_path()
#         Parent_Folder_Path = BASE_DIR
        
#         output_filepath = Parent_Folder_Path / "flickr_photos_with_sentiment.xlsx"
#         top_10_images_filepath = Parent_Folder_Path / "top_10_images.xlsx"
        
#         print(output_filepath)
        
        
        
        
#         objExcelOperations = ExcelOperations() 
         
#         error_status =  objFlicker_Data_Download.get_flickr_data( api_key, secret_api_key , output_filepath , latitude , longitude ,pagecnt ) 

#         if (error_status == 'Error Occured'): 
#             self.display_text('>>> Error occured while executing WanderSphere AI Tour Geoguide: A Real-Time Sentiment-Driven Travel Recommendation System' + '\n'  , 'Error')   
              
#             self.display_text('>>> Terminating Execution' + '\n'  , 'Error')  
#             self.text.config(state='disabled') 
#             raise NameError('EXCEL Operation Exception')     
#             stopcodehere       
         
        
#         #--(Step - 2: read and perform data processing on the flicker dataset)----
        
#         error_status = objExcelOperations.read_flickrdata(output_filepath ,  Parent_Folder_Path )
#         print(error_status)
        

#         if (error_status == 'Error Occured'):
 
            
#             self.display_text('>>> Error occured while executing WanderSphere AI Tour Geoguide: A Real-Time Sentiment-Driven Travel Recommendation System' + '\n'  , 'Error')  
            
#             self.display_text('>>> Terminating Execution' + '\n'  , 'Error')  
#             self.text.config(state='disabled') 
#             raise NameError('EXCEL Operation Exception')     
#             stopcodehere       
         
            
#         ######### Define the images and their URLs along with Google Maps links
#         images_info = objExcelOperations.dataprocessing(top_10_images_filepath)
        
        
#         ######################  Display Result GUI ##############
#         print(images_info)
#         self.open_new_window(images_info) 
  
#         self.display_text('>>> Project execution completed successfully   '  + '\n\n'  , 'Normal')
#         time.sleep(1)   
#         try:
#             os.remove(top_10_images_filepath)
#             os.remove(output_filepath) 
#         except:pass 
            
#         self.text.config(state='disabled') 
     

# time.sleep(1)     
# root = tkinterApp()


# ############################################################################
 
# #root = tk.Tk()
# root.title("WanderSphere AI Tour Geoguide: A Real-Time Sentiment-Driven Travel Recommendation System (Version: 1.0.0)")
# root.geometry('1393x769')

# root.pack_propagate(0)
# root.resizable(0,0)

# def OnFocusIn(event):
#     if type(event.widget).__name__ == 'WanderSphere AI Tour Geoguide: A Real-Time Sentiment-Driven Travel Recommendation System (Version: 1.0.0)':
#         event.widget.attributes('-topmost', False)

# root.attributes('-topmost', True)
# root.focus_force()
# root.bind('<FocusIn>', OnFocusIn)

# root.mainloop()

     


