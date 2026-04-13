# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr 10 22:57:56 2026

@author: komalwavhal
"""
 

"""
main.py

Tkinter launcher for the Trade Reconciliation project.


- Outputs are written under ./output
"""


#try:
#    import docx
#except:
#    os.system('python -m pip install -i python-docx')
#    import docx
   
 
import os
import tkinter as tk
from tkinter import ttk
from PIL import Image,ImageTk
import threading
import time
from tkinter import scrolledtext
from tkinter import messagebox
LARGEFONT =("Verdana", 35) 
import requests
from io import BytesIO
import webbrowser
from urllib.parse import quote_plus
import pandas as pd 

class tkinterApp(tk.Tk):
    
    def get_base_path(self):
        from pathlib import Path
        return Path(__file__).resolve().parent
        
    # __init__ function for class tkinterApp   
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, Page1):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)
       
    
    
    # to display the current frame passed as  parameter
    def show_frame(self, cont): 
    
        frame = self.frames[cont]
        frame.tkraise() 
            
        BASE_DIR = self.get_base_path()
        print(' system BASE_DIR -  ' , BASE_DIR)
        


# first window frame startpage 
class StartPage(tk.Frame):

    
    def display_text(self, text, tag):
        def _append():
            self.text.tag_config('Error', background="white", foreground="red", font=('calibri', 12))
            self.text.tag_config('Normal', background="white", foreground="#819ec4", font=('calibri', 12))
            self.text.insert("end", text, tag)
            self.text.see("end")
    
        self.text.after(0, _append)

    def __init__(self, parent, controller): 
            

        tk.Frame.__init__(self, parent)

        BASE_DIR = controller.get_base_path()
        print("system BASE_DIR -", BASE_DIR)

        Img_Path = BASE_DIR / "Images"
        imgFilePath = Img_Path / "pic-11.png"
        AFFIRM_ICON_imgFilePath = Img_Path / "AFFIRM_ICON.png"

        img = Image.open(imgFilePath)
        photo = ImageTk.PhotoImage(img)

        lab1 = ttk.Label(self, image=photo)
        lab1.place(x=0)
        lab1.image = photo

        affirm_Button_image = Image.open(AFFIRM_ICON_imgFilePath)
        photo2 = ImageTk.PhotoImage(affirm_Button_image)

        affirm_Button = tk.Button(
            self,
            image=photo2,
            command=lambda: controller.show_frame(Page1)
        )
        affirm_Button.place(x=577, y=420)
        affirm_Button.image = photo2

        self.configure(background='#FFFFFF')
        
        
        ######### image path   ###########
        from pathlib import Path
        
        BASE_DIR = Path(__file__).resolve().parent
        Img_Path = BASE_DIR / "Images"
        
        # userManual_Button_imgFilePath = Img_Path / "UserManual.png"
        imgFilePath = Img_Path / "pic-11.png"
        AFFIRM_ICON_imgFilePath = Img_Path / "AFFIRM_ICON.png"
        
        # print("BASE_DIR:", BASE_DIR)
        # print("Img_Path:", Img_Path)
        # print("Image exists:", imgFilePath.exists())
        


# second window frame page1
class Page1(tk.Frame): 
        
    def browsefunc_1(self):
        from pathlib import Path
        import os
        import platform
        import subprocess
    
        BASE_DIR = Path(__file__).resolve().parent
        User_Manual_Path = BASE_DIR / "User_Manual"
    
        # print("BASE_DIR:", BASE_DIR)
        # print("User_Manual_Path:", User_Manual_Path)
        # print("Exists:", User_Manual_Path.exists())
    
        if not User_Manual_Path.exists():
            messagebox.showerror("Error", f"Folder not found:\n{User_Manual_Path}")
            return
    
        system_name = platform.system()
    
        if system_name == "Windows":
            os.startfile(User_Manual_Path)
        elif system_name == "Darwin":  # macOS
            subprocess.run(["open", str(User_Manual_Path)])
        else:  # Linux
            subprocess.run(["xdg-open", str(User_Manual_Path)])
            
              
        
     ########################################################################################################
     ########################################################################################################

    def __init__(self, parent, controller): 
        
        BASE_DIR = controller.get_base_path()
        # print("system BASE_DIR -", BASE_DIR)
         
        Img_Path = BASE_DIR / "Images"
        
        imgFilePath = Img_Path / "pic-1.png"
        Bck_Button_imgFilePath = Img_Path / "Bck_Button.png"
        run_Button_imgFilePath = Img_Path / "run.PNG"
        
        # print("Img_Path:", Img_Path)
        # print("pic-1 exists:", imgFilePath.exists())
        # print("Back button exists:", Bck_Button_imgFilePath.exists())
        # print("Run button exists:", run_Button_imgFilePath.exists())
        
        ##------------------------------------------------------------------------------------------------------------
          
        ###-----------------------------------------------------------------------------------------------------------

        		
        tk.Frame.__init__(self, parent)
        
        BASE_DIR = controller.get_base_path()
        # print("system BASE_DIR -", BASE_DIR)
         
        Img_Path = BASE_DIR / "Images"
        
        imgFilePath = Img_Path / "pic-1.png"
        Bck_Button_imgFilePath = Img_Path / "Bck_Button.png"
        run_Button_imgFilePath = Img_Path / "run.PNG"
        
        # print("Img_Path:", Img_Path)
        # print("pic-1 exists:", imgFilePath.exists())
        
        img = Image.open(imgFilePath)
        
        photo = ImageTk.PhotoImage(img)
        lab1 = ttk.Label(self, image=photo)
        lab1.pack()
        lab1.place(x=0)
        lab1.image = photo
        
        Bck_Button_image = Image.open(Bck_Button_imgFilePath)
        photo = ImageTk.PhotoImage(Bck_Button_image)
        Bck_Button = tk.Button(self, image=photo, command=lambda: controller.show_frame(StartPage))
        Bck_Button.place(x=37, y=30)
        Bck_Button.image = photo
        
        self.configure(background='#FFFFFF')
                
        import getpass
        import platform
        
        user = getpass.getuser()
        machine_name = platform.node()
        fullname = user
        
        # print("user:", user)
        # print("machine_name:", machine_name)
        # print("fullname:", fullname)
        
        
        #### ----------###(Application Logs)---------------------------------- 
        self.text = scrolledtext.ScrolledText(self, wrap='word', bg='#FFFFFF', height=7, width=96, font=('calibri',12), relief='ridge')
        # self.text = scrolledtext.ScrolledText(self,wrap = WORD,bg = '#FFFFFF',  height=3, width=46,  font=('calibri',12),relief= 'ridge') 
        self.text.place(x=200,y=500)
        
        self.configure(background='#FFFFFF') 
           
        ###-----------------------------------------------------------------------------------------
        
        
        # Internal trades file
        self.internal_file_var = tk.StringVar()
        self.market_file_var = tk.StringVar()
        
        tk.Label(
            self,
            text="Internal Trades CSV:",
            bg="#FFFFFF",
            fg="black",
            font=("calibri", 12, "bold")
        ).place(x=180, y=180)
        
        self.internal_file_entry = tk.Entry(
            self,
            textvariable=self.internal_file_var,
            width=65,
            font=("calibri", 11)
        )
        self.internal_file_entry.place(x=360, y=180)
        
        tk.Button(
            self,
            text="Browse",
            command=self.browse_internal_file,
            font=("calibri", 11, "bold"),
            bg="#D9EAF7"
        ).place(x=930, y=176)
        
        
        # Market prices file
        tk.Label(
            self,
            text="Market Prices CSV:",
            bg="#FFFFFF",
            fg="black",
            font=("calibri", 12, "bold")
        ).place(x=180, y=230)
        
        self.market_file_entry = tk.Entry(
            self,
            textvariable=self.market_file_var,
            width=65,
            font=("calibri", 11)
        )
        self.market_file_entry.place(x=360, y=230)
        
        tk.Button(
            self,
            text="Browse",
            command=self.browse_market_file,
            font=("calibri", 11, "bold"),
            bg="#D9EAF7"
        ).place(x=930, y=226)
        
        
         
        #####RUN Automatoin  
        run_image = Image.open(run_Button_imgFilePath) 
        photo = ImageTk.PhotoImage(run_image)        
        button_Export = tk.Button(self,image=photo,command=  self.Execute_speaker_eng)  ###,bg = "white", bd = 0
        button_Export.place(x=475,y=310)        
        button_Export.image = photo  
                  
        ##### ----------###(THIS STYLE FOR THE PROGRESSBAR)----------------------------------
        self.style = ttk.Style(self)
        
        self.style.layout('text.Horizontal.TProgressbar',
        	 [('Horizontal.Progressbar.trough',
        	   {'children': [('Horizontal.Progressbar.pbar',
        					  {'side': 'left', 'sticky': 'ns'})],
        		'sticky': 'nswe'}),
        	  ('Horizontal.Progressbar.label', {'sticky': ''})])   
        self.style.configure('text.Horizontal.TProgressbar', text=' ')
        
        
        self.configure(background='#FFFFFF')     

    
    def browse_internal_file(self):
        
        from tkinter import filedialog
        root_window = self.winfo_toplevel()
        root_window.iconify()  # minimize app

        
        selected_file = filedialog.askopenfilename(
            title="Select Internal Trades CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        root_window.deiconify()  # restore app
        root_window.lift()
        root_window.focus_force()
        
        
        if selected_file:
            self.internal_file_var.set(selected_file)
    
    def browse_market_file(self):
        
        from tkinter import filedialog
        root_window = self.winfo_toplevel()
        root_window.iconify()  # minimize app

        
        selected_file = filedialog.askopenfilename(
            title="Select Market Prices CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
            
        root_window.deiconify()  # restore app
        root_window.lift()
        root_window.focus_force()
        
        
        if selected_file:
            self.market_file_var.set(selected_file)
            
            
    def Execute_speaker_eng(self) :  
        # run process in a thread to avoid blocking gui 
        t = threading.Thread(target=self.execute_main)
        t.start()
 
    
    # Function to load and display the image
    def display_image(self, image_name, image_url, google_map_link, row, col, checkbuttons):
        try:
            # Fetch image from the URL
            response = requests.get(image_url)
            if response.status_code == 200:
                # Open the image using Pillow
                img_data = BytesIO(response.content)
                img = Image.open(img_data)
                
                # Resize image to fit within the window (optional)
                img = img.resize((200, 200))  # Adjusted size for 5 images per row
                
                # Convert the image for Tkinter compatibility
                img_tk = ImageTk.PhotoImage(img)
                
                # Create labels to display the Google Map link (above the image)
                def open_map():
                    webbrowser.open(google_map_link)
    
                label_map = tk.Label(self.new_window, text="View on Google Maps", fg="blue", bg='#FFFFFF', cursor="hand2", font=("calibri", 15))
                label_map.grid(row=row*4, column=col, padx=10, pady=5)
                label_map.bind("<Button-1>", lambda e: open_map())  # Bind the click event to open the map link
                
                # Create the image label
                label_image = tk.Label(self.new_window, image=img_tk)
                label_image.image = img_tk  # Keep a reference to avoid garbage collection
                label_image.grid(row=row*4+1, column=col, padx=10, pady=10)
                
                # Create the image name label (below the image)
                label_name = tk.Label(self.new_window, text=image_name, fg="black", bg='#FFFFFF', font=("calibri", 12), wraplength=200)
                label_name.grid(row=row*4+2, column=col, padx=10, pady=10)
    
                # Create checkbox for the image (below the image name)
                var = tk.BooleanVar()
                checkbox = tk.Checkbutton(self.new_window, fg="black", bg='#FFFFFF', text=" ", variable=var)
                checkbox.grid(row=row*4+3, column=col, padx=10, pady=5)
    
                # Store the checkbox state and the place name in the list
                checkbuttons.append((var, image_name))  # Store place_name (not link)
    
                self.new_window.configure(background='#FFFFFF') 
    
            else:
                print(f"Error: Unable to fetch image {image_name}")
        except Exception as e:
            print(f"Failed to load image {image_name}: {str(e)}")
    
    # Function to generate the Google Map itinerary based on selected checkboxes
    def generate_itinerary(self, checkbuttons):
        # Collect the selected locations (place names) based on the checkbox state
        selected_locations = [place_name for var, place_name in checkbuttons if var.get()]
    
        if not selected_locations:
            messagebox.showinfo("No Selection", "Please select at least one place to generate the itinerary.")
            return
    
        # Create Google Maps URL with the selected locations as stops
        base_url = "https://www.google.com/maps/dir/"
        destination_url = base_url + "/".join([quote_plus(location) for location in selected_locations])
    
        # Open the generated itinerary in the browser
        webbrowser.open(destination_url)

        
    def display_text(self, text, tag):
        def _append():
            self.text.tag_config('Error', background="white", foreground="red", font=('calibri', 12))
            self.text.tag_config('Normal', background="white", foreground="#819ec4", font=('calibri', 12))
            self.text.insert("end", text, tag)
            self.text.see("end")
    
        self.text.after(0, _append)
    
    def open_new_window(self, images_info):
        
        # Create a new top-level window
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Image Display with Google Map Itinerary")
        
        # Set the size of the new window
        # self.new_window.geometry("300x200")
         
        # List to store checkbox variables and their corresponding place names (not links)
        checkbuttons = []
        
        # Display all the images with their names, Google Maps links, and checkboxes
        for index, (image_name, image_url, google_map_link) in enumerate(images_info):
            row = index // 5  # Determine the row number (5 images per row)
            col = index % 5   # Determine the column number (5 images per row)
            self.display_image(image_name, image_url, google_map_link, row, col, checkbuttons)
        
        # Add a "Generate Google Map Itinerary" button to generate a single itinerary
        generate_button = tk.Button(self.new_window, text="Generate Google Map Itinerary",fg= 'Black', bg = '#FFFFFF',font=('calibri',15,'bold'), command=lambda: self.generate_itinerary(checkbuttons))
        generate_button.grid(row=(len(images_info)//5)*4+1, column=0, columnspan=5, pady=20)
        
        # Start the Tkinter event loop
        # self.new_window.mainloop()
 
        
    def execute_main(self):  
       
        self.text.config(state='normal')
        self.text.delete("1.0" , "end")
        
        import datetime
        now = datetime.datetime.now() 
        x = str(now.strftime("%Y-%m-%d %H:%M:%S"))  
        self.display_text('>>> Project execution started at ' + x + '\n\n'  , 'Normal')
        time.sleep(1)   
         
        ########################################################################################
              
        error_status = ''
        
        ##### Importing Custom Libraries for Excel Operations
        from pathlib import Path
        from reconciliation_engine import Reconciliation, ReconConfig
                
        BASE_DIR = Path(__file__).resolve().parent
        internal_file_path = self.internal_file_var.get().strip()
        
        
        market_file_path = self.market_file_var.get().strip()
        
        if not internal_file_path: 
                
            self.display_text('>>> Please browse and select Internal Trades CSV' +  '\n\n'  , 'Error')
            
            return
        
        if not market_file_path:
            self.display_text(">>> Please browse and select Market Prices CSV\n", "Error")
            return
        
        config = ReconConfig(
            internal_file=Path(internal_file_path),
            market_file=Path(market_file_path),
            output_file=BASE_DIR / "output" / "reconciliation_detail.csv",
            exceptions_file=BASE_DIR / "output" / "exceptions_detail.csv",
            metadata_file=BASE_DIR / "output" / "run_metadata.json",
            abs_tolerance=1.0,
            pct_tolerance=1.0,
        )

         
        try:
            engine = Reconciliation(config)
            result = engine.run()
        
            print(result["matched"])
            print(result["warnings"])
        
        except Exception as exc:
            print(f"error occoured: {exc}")
            self.display_text(f'>>> Error details: {exc}\n', 'Error')
            error_status = 'Error Occured'
            
            
        ############################################################################################################################################################################### 
         
        
        ##### Build Tableau Output  
        try:
            input_file = BASE_DIR / "output" / "reconciliation_detail.csv"
            df = pd.read_csv(input_file)
            


            # KPI Summary
            kpi = pd.DataFrame({
                "total_trades": [len(df)],
                "matched": [(df["recon_status"] == "MATCHED").sum()],
                "warnings": [(df["recon_status"] == "TOLERANCE_WARNING").sum()],
                "breaks": [(df["recon_status"] == "PRICE_BREAK").sum()],
                "missing": [(df["recon_status"] == "MARKET_DATA_MISSING").sum()]
            })
    
            kpi["match_rate"] = (kpi["matched"] / kpi["total_trades"]) * 100
     
    
            # Status Summary
            status_summary = df["recon_status"].value_counts().reset_index()
            status_summary.columns = ["status", "count"]
     
            # Severity Summary
            severity_summary = df["severity"].value_counts().reset_index()
            severity_summary.columns = ["severity", "count"]
     
    
            # Symbol Breakdown
            symbol_breaks = df[df["recon_status"] == "PRICE_BREAK"] \
                .groupby("symbol") \
                .size() \
                .reset_index(name="break_count") \
                .sort_values(by="break_count", ascending=False)
     
        
            kpi.to_csv(BASE_DIR / "output" / "kpi_summary.csv", index=False)
            status_summary.to_csv(BASE_DIR / "output" / "status_summary.csv", index=False)
            severity_summary.to_csv(BASE_DIR / "output" / "severity_summary.csv", index=False)
            symbol_breaks.to_csv(BASE_DIR / "output" / "symbol_breaks.csv", index=False)
            
    
            print("✅ Tableau files generated successfully")
            
        
        except Exception as exc:
            print(f"error occoured: {exc}")
            self.display_text(f'>>> Error details: {exc}\n', 'Error')
            error_status = 'Error Occured'
            
         
        

        if (error_status == 'Error Occured'): 
            self.display_text('>>> Error occured while executing trade reconciliation automation system\n', 'Error')
            self.display_text('>>> Terminating Execution\n', 'Error')
            self.text.config(state='disabled')
            stopcodehere
                 
        
        ######################  Email Notification ##############
                
        ##### Send Email
        try:
            from Email_Utils import EmailOperations
        
            emailer = EmailOperations(
                to_email="kwavhal@stevens.edu",
                sender_name="Komal",
                base_dir=str(BASE_DIR)
            )
        
            emailer.send_email()
            self.display_text(">>> Email sent successfully with attached output files\n", "Normal")
        
        except Exception as exc:
            print(f"error occoured while sending email: {exc}")
            self.display_text(f'>>> Email Error details: {exc}\n', 'Error')
            self.display_text('>>> Terminating Execution\n', 'Error')
            self.text.config(state='disabled')
            stopcodehere
             
        self.display_text(">>> Project execution completed successfully  ✅ Tableau files generated successfully\n", "Normal")
            
        self.text.config(state='disabled') 
     

time.sleep(1)     
root = tkinterApp()


############################################################################
 
#root = tk.Tk()
root.title(" Trade Reconciliation & Exception Management Platform (Version: 1.0.0)")
root.geometry('1203x719')

root.pack_propagate(0)
root.resizable(0,0)

def OnFocusIn(event):
    if type(event.widget).__name__ == 'Trade Reconciliation & Exception Management Platform (Version: 1.0.0)':
        event.widget.attributes('-topmost', False)

root.attributes('-topmost', True)
root.focus_force()
root.bind('<FocusIn>', OnFocusIn)

root.mainloop()

     


