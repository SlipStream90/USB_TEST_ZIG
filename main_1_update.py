from turtle import color
import customtkinter as ctk
import time
import threading
from packets import send_receive_data , packets, identifier, USB_number, voltage_check,list_mice_devices,send_data,vcc_led_check
from com_port import list_active_ft232r_devices
import serial as pyserial
from tkinter import messagebox
import csv
from colorama import Fore, init
import win32com.client
import re
import time


def data_store(led_max, led_min, vcc_min, vcc_max):
    try:
        led_set = [led_max, led_min]
        vcc_set = [vcc_max, vcc_min]
        main=open('settings.csv',"w")
        write=csv.writer(main)
        write.writerows([vcc_set,led_set])
        main.close()
    except Exception as e:
        messagebox.showerror("Error", e)


def open_win():
    new = ctk.CTkToplevel(root)
    new.geometry('640x480')
    new.title("Enter Password")
    new.resizable(False, False)
    new.protocol("WM_DELETE_WINDOW", lambda: on_closing(new))

    title_label = ctk.CTkLabel(new, text="Enter Your Password", font=("Arial", 18, "bold"),text_color="white",corner_radius=5)
    title_label.pack(pady=(40, 20))

    password_entry = ctk.CTkEntry(new, show="*", width=300, font=("Arial", 14))
    password_entry.pack(pady=100)

    submit_button = ctk.CTkButton(new, text="Submit", command=lambda: validate_password(new, password_entry.get()))
    submit_button.pack(pady=20)

    new.bind("<Return>", lambda event: validate_password(new, password_entry.get()))




def validate_password(window, password):
    if password == "yourpassword":
        display_sliders()
        window.destroy()

def on_closing(window):
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        window.destroy()

def display_sliders():
    new_window = ctk.CTkToplevel(root)
    new_window.geometry('640x480')
    new_window.title("VCC and LED Voltage")
    new_window.resizable(False, False)
    try:
        main=open("settings.csv","r")
        store=[]
        read=csv.reader(main)
        for i in read:
            if i==[]:
                pass
            else:
                store.append(i)
    except Exception as e:
        messagebox.showerror(e)

    vcc_label = ctk.CTkLabel(new_window, text="VCC(min,max)", font=("Arial", 16, "bold"),text_color="white",corner_radius=5)
    vcc_label.grid(row=0, column=0, padx=(40, 10), pady=(40, 10))
    
    vcc_min_entry = ctk.CTkEntry(new_window, width=100)
    vcc_min_entry.grid(row=0, column=1, padx=(10, 40), pady=(40, 10))
    vcc_min_entry.insert(0, store[0][1])  # Set min value

    vcc_max_entry = ctk.CTkEntry(new_window, width=100)
    vcc_max_entry.grid(row=0, column=2, padx=(10, 40), pady=(40, 10))
    vcc_max_entry.insert(0, store[0][0])  # Set max value

    led_voltage_label = ctk.CTkLabel(new_window, text="LED Voltage(min,max)", font=("Arial", 16, "bold"),text_color="white",corner_radius=5)
    led_voltage_label.grid(row=1, column=0, padx=(40, 10), pady=(40, 10))

    led_min_entry = ctk.CTkEntry(new_window, width=100)
    led_min_entry.grid(row=1, column=1, padx=(10, 40), pady=(40, 10))
    led_min_entry.insert(0, store[1][1])  # Set min value

    led_max_entry = ctk.CTkEntry(new_window, width=100)
    led_max_entry.grid(row=1, column=2, padx=(10, 40), pady=(40, 10))
    led_max_entry.insert(0, store[1][0])  # Set max value
    

    def fetcher():
        global led_max, led_min, vcc_max, vcc_min
        led_max = float(led_max_entry.get())
        led_min = float(led_min_entry.get())
        vcc_max = float(vcc_max_entry.get())
        vcc_min = float(vcc_min_entry.get())

        if led_max <= led_min or vcc_max <= vcc_min:
            messagebox.showerror("Error", "Max value should be greater than min value")
        else:

            confirm_window = ctk.CTkToplevel(new_window)
            confirm_window.geometry('300x200')
            confirm_window.title("Confirm Changes")
            confirm_window.resizable(False, False)

            confirm_label = ctk.CTkLabel(confirm_window, text="Are you sure you want to change the settings?", 
                                     font=("Arial", 14), wraplength=250)
            confirm_label.pack(pady=(20, 10))

            def confirm_changes():
                data_store(led_max, led_min, vcc_min, vcc_max)
                confirm_window.destroy()
                new_window.destroy()
                messagebox.showinfo("Success", "Settings have been updated successfully.")

            def cancel_changes():
                confirm_window.destroy()

            confirm_button = ctk.CTkButton(confirm_window, text="Confirm", command=confirm_changes)
            confirm_button.pack(pady=(10, 5))

            cancel_button = ctk.CTkButton(confirm_window, text="Cancel", command=cancel_changes)
            cancel_button.pack(pady=5)

    submit_button = ctk.CTkButton(new_window, text="Submit", command=fetcher)
    submit_button.grid(row=2, column=1, padx=(10, 40), pady=(40, 10))
    new_window.protocol("WM_DELETE_WINDOW", lambda: on_closing(new_window))


class CircularStatus(ctk.CTkCanvas):
    def __init__(self, parent, size=200, fg_color="#008080", bg_color="#E0E0E0", **kwargs):
        super().__init__(parent, width=size, height=size, highlightthickness=0, **kwargs)
        self.size = size
        self.fg_color = fg_color
        self.bg_color = bg_color
        self.extent = 0
        self.is_animating = False
        self.draw()

    def draw(self):
        self.delete("all")
        self.create_oval(10, 10, self.size-10, self.size-10, fill=self.bg_color, outline="")
        self.arc = self.create_arc(10, 10, self.size-10, self.size-10, start=90, extent=self.extent,
                                   fill=self.fg_color, outline="")
        self.create_oval(self.size/4, self.size/4, 3*self.size/4, 3*self.size/4,
                         fill="white", outline="")
        self.text = self.create_text(self.size/2, self.size/2, text="",
                                     font=("Arial", int(self.size/10), "bold"), fill="black")

    def animate(self, status):
        if self.is_animating:
            return
        
        self.is_animating = True
        threading.Thread(target=self._animate, args=(status,), daemon=True).start()

    def _animate(self, status):
        if status.upper() == "PASS":
            self.fg_color = "#4CAF50"  # Green
            final_text = "PASS"
        elif status.upper() == "FAIL":
            self.fg_color = "#F44336"  # Red
            final_text = "FAIL"
        else:
            self.fg_color = "#F44336"  # Yellow for unknown status
            final_text = "START"

        self.extent = 0
        while self.extent < 360:
            self.extent += 10
            self.itemconfig(self.arc, extent=self.extent, fill=self.fg_color)
            self.update()
            time.sleep(0.03)

        self.itemconfig(self.text, text=final_text)
        self.is_animating = False

class Ziggui:
    def __init__(self, root):
        self.root = root
        self.root.title("Hub Zig")
        self.root.resizable(False, False)

        self.root.geometry("1200x850")

        # Set custom color theme
        ctk.set_default_color_theme("blue")  # Start with a base theme
        ctk.ThemeManager.theme["CTkButton"]["fg_color"] = ["#1A2B3C", "#2C3E50"]
        ctk.ThemeManager.theme["CTkButton"]["hover_color"] = ["#2C3E50", "#34495E"]
        ctk.ThemeManager.theme["CTkCheckBox"]["fg_color"] = ["#008080", "#00A3A3"]
        ctk.ThemeManager.theme["CTkEntry"]["fg_color"] = ["#F0F2F5", "#E0E0E0"]
        ctk.ThemeManager.theme["CTkFrame"]["fg_color"] = ["#FFFFFF", "#1A2B3C"]
        ctk.ThemeManager.theme["CTkLabel"]["fg_color"] = ["#1A2B3C", "#ECF0F1"]

        try:
            led_set = ["6", "4"]
            vcc_set = ["6", "4"]
            main=open('settings.csv',"w")
            write=csv.writer(main)
            write.writerows([vcc_set,led_set])
            main.close()
        except Exception as e:
            messagebox.showerror("Error", e)


        self.theme_var = ctk.StringVar(value="light")
        ctk.set_appearance_mode(self.theme_var.get())

        # Set custom fonts
        self.header_font = ("Roboto", 24, "bold")
        self.body_font = ("Open Sans", 14)

        # Create and pack widgets with refined styling
        self.create_header()
        self.create_content()
        self.create_table()

    def create_header(self):
        self.header_frame = ctk.CTkFrame(self.root, corner_radius=0)
        self.header_frame.pack(fill="x", pady=(0, 20))

        self.main_label = ctk.CTkLabel(self.header_frame, text="USB HUB ZIG", font=self.header_font,text_color="white",corner_radius=5)
        self.main_label.pack(side="left", padx=20, pady=10)

    def create_content(self):
        self.content_frame = ctk.CTkFrame(self.root)
        self.content_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.widget_frame = ctk.CTkFrame(self.content_frame, width=300, corner_radius=10)
        self.widget_frame.pack(side="left", padx=(0, 20), fill="y")

        self.admin_call = ctk.CTkButton(self.widget_frame, text="Admin Login", font=self.body_font,
                                        corner_radius=8, height=40, command=open_win)
        self.admin_call.pack(pady=60)

        self.start_button = ctk.CTkButton(self.content_frame, text="Start Test", font=self.body_font,
                                          corner_radius=8, height=50, command=self.toggle_status)
        self.start_button.pack(pady=(0, 20))

        # Add status widget with drop shadow
        self.status_widget = CircularStatus(parent=self.content_frame, size=200)
        self.status_widget.pack(side="left", padx=(0, 20), pady=(0, 20))        


    def list_mice_devices(self):
        wmi = win32com.client.GetObject("winmgmts:")
        devices = wmi.InstancesOf("Win32_PointingDevice")
        mice_list = []

        for device in devices:
            name = device.Name
            match = re.search(r'VID_(\w+)&PID_(\w+)', device.PNPDeviceID)
            vid, pid = match.groups() if match else (None, None)
            mice_list.append((name, vid, pid))     


    def check_devices(self):
        com_ports = list_active_ft232r_devices()
        if not com_ports:
            self.clear_table()
            print("cleared")  # Clear the table if no devices are connected 
        
        self.identifier=self.root.after(100, self.check_devices)
        return com_ports



    def clear_table(self):
        for cell in self.table_cells:
            cell.configure(text="")
        print("Table cleared due to disconnection.")

    def center_window(self, root, width, height):
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry("1200x850")



    def disable_button(self):
        self.start_button.configure(state="disabled", fg_color="gray", text="Processing...")
        self.root.update_idletasks()

    def enable_button(self):
        self.start_button.configure(state="normal", fg_color=["#3a7ebf", "#1f538d"], text="Start Test")        


    def create_table(self):

        self.table_frame = ctk.CTkFrame(master=self.content_frame, corner_radius=10)
        self.table_frame.pack(fill="both", expand=True, pady=(20, 0))

        com_ports = list_active_ft232r_devices()
        while len(com_ports) < 6:
            com_ports.append("-")
        data = [["Test Details", "USB Port/Voltage", "Status","Max Value","Min Value"]]
        rows=["VCC", "LED", "USB1", "USB2", "USB3", "USB4"]
        for i, port in enumerate(com_ports):
            data.append([f"{rows[0]}", "", "","N/A","N/A"])
            rows.pop(0)

        self.table_rows = len(data)
        self.table_cols = len(data[0])
        self.table_cells = []

        table_border = ctk.CTkFrame(self.table_frame, fg_color="gray")
        table_border.pack(padx=1, pady=1, expand=True, fill="both")

        for i in range(self.table_rows):
            for j in range(self.table_cols):
                global cell_frame
                cell_frame = ctk.CTkFrame(master=table_border, corner_radius=0, fg_color="white", border_width=1, border_color="gray")
                cell_frame.grid(row=i, column=j, sticky="nsew", padx=1, pady=1)
                
                cell_label = ctk.CTkLabel(master=cell_frame, text=data[i][j], font=("Arial", 14), fg_color="transparent")
                cell_label.pack(expand=True, fill="both", padx=5, pady=5)
                
                if i == 0:  # Header row
                    cell_frame.configure(fg_color="lightgray")
                    cell_label.configure(font=("Arial", 14, "bold"))
                
                self.table_cells.append(cell_label)

        for i in range(self.table_rows):
            table_border.grid_rowconfigure(i, weight=1)
        for j in range(self.table_cols):
            table_border.grid_columnconfigure(j, weight=1)
     
    def toggle_status(self):

        self.disable_button()  # Disable button while processing
        com_ports = self.check_devices()
        if not com_ports:
            self.clear_table()  # Clear the table if no COM ports are available
            messagebox.showinfo("Error","No Ports detected. Please check the connection.")
            self.enable_button()  # Re-enable the button
            return
        data = [["Test Description", "USB Port/Voltage", "Status","Max Value","Min Value"]]
        DEL=["USB1","USB2","USB3","USB4"]
        port=list_active_ft232r_devices()[0]
        print(port)
        vcc_check_done=False
        led_check_done=False
        test_results = []
        try:
            ser = pyserial.Serial(port, 9600, timeout=2, write_timeout=2)
            time.sleep(0.1)
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            pack=packets()
            pack.pop(0)
            for i in pack:
                n=pack.index(i)
                if send_data(ser,i):
                    test_port= list_mice_devices()
                    time.sleep(0.1)
                    usb_port=[tup for tup in test_port if "USB Input Device" in tup] #add the respective pid for the IC 
                    if len(usb_port)==0:
                        test_results.append([f"Test for USB{n+1}",f"USB{n+1}" ,"FAIL","N/A","N/A"])
                        self.status_widget.animate("FAIL")
                        time.sleep(0.2)       
                    else:
                        if len(usb_port)==2:
                            usb_port.pop(1)
                        volt=voltage_check(ser,vcc_led_check()[1],active_device=port)
                        try:
                            main=open("settings.csv","r")
                            store=[]
                            read=csv.reader(main)
                            for i in read:
                                if i==[]:
                                    pass
                                else:
                                    store.append(i)
                        except Exception as e:
                            messagebox.showerror("error",e)           

                        if not led_check_done:
                            led_volt=voltage_check(ser,vcc_led_check()[2],active_device=port)
                            if float(store[1][1])<=float(led_volt) and float(store[1][0])>=float(led_volt):
                                test_results.append(["LED Voltage",led_volt,"PASS",store[1][0],store[1][1]])
                            else:
                                test_results.append(["LED Check",led_volt,"FAIL",store[1][0],store[1][1]])
                                self.status_widget.animate("FAIL")
                            led_check_done=True

                        if not vcc_check_done:
                            if float(store[0][1])<=float(volt) and float(store[0][0])>=float(volt):
                                test_results.append(["VCC Check",volt,"PASS",store[0][0],store[0][1]])
                            else:
                                test_results.append(["VCC Check",volt,"FAIL",store[0][0],store[0][1]])
                                self.status_widget.animate("FAIL")
                            vcc_check_done=True           
                        test_results.append([f"Test for USB{n+1}",f"USB{n+1}" , "PASS","N/A","N/A"])
                        time.sleep(0.2)# time to be edited to minimize gap between switching
        except pyserial.SerialException as e:
            test_results.append(["-", "-", "FAIL","N/A","N/A"])
            print(f"Error opening or communicating with port {port}: {e}")
        finally:
            if 'ser' in locals() and ser.is_open:
                ser.close()
                print(f"Closed port {port}")

    # Assign USB ports to remaining results and sort
        for result in test_results:
            if result[0] == "-":
                result[0] = f"Test for {DEL.pop(0)}"
    
    # Sort results in the specified order
        sorted_results = []
        for test_type in ["LED", "VCC", "USB1", "USB2", "USB3", "USB4"]:
            result = next((r for r in test_results if test_type in r[0]), [f"Test for {test_type}", "-", "FAIL","N/A","N/A"])
            sorted_results.append(result)
        data.extend(sorted_results)

    # Update the table
        for i in range(self.table_rows):
            for j in range(self.table_cols):
                if i < len(data) and j < len(data[i]):
                    cell = self.table_cells[i * self.table_cols + j]
                    cell.configure(text=data[i][j])
                    if j == 2 and data[i][j] == "FAIL":  # Status column
                        cell.master.configure(fg_color="red")
                    elif i > 0:  # Not header row
                        cell.master.configure(fg_color="white")
                else:
                    self.table_cells[i * self.table_cols + j].configure(text="")

    # Determine overall status
        if all(result[2] == "PASS" for result in sorted_results if result[2] != "Not Tested"):
            self.status_widget.animate("PASS")
        else:
            self.status_widget.animate("FAIL")

        self.root.after(500, self.enable_button)


    def clear_table(self,):
        for i in range(self.table_rows):
            for j in range(self.table_cols):
                if i > 0 and j > 0:  # Skip first row and first column
                    cell_index = i * self.table_cols + j
                    cell = self.table_cells[cell_index]
                    cell.configure(text="")
                    cell.master.configure(fg_color="white")
                    n=0

    def toggle_theme(self):
        ctk.set_appearance_mode(self.theme_var.get())





if __name__ == "__main__":

    root = ctk.CTk()
    app = Ziggui(root)
    root.mainloop()