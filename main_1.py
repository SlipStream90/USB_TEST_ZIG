
import customtkinter as ctk
import time
import threading
from packets import send_receive_data, crc_16bit_modbus, packets, main, identifier, USB_number, voltage_check
from com_port import list_active_ft232r_devices
import serial as pyserial
from tkinter import messagebox
import csv


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

    title_label = ctk.CTkLabel(new, text="Enter Your Password", font=("Arial", 18, "bold"))
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
    main=open("settings.csv","r")
    store=[]
    read=csv.reader(main)
    for i in read:
        if i==[]:
            pass
        else:
            store.append(i)

    vcc_label = ctk.CTkLabel(new_window, text="VCC(min,max)", font=("Arial", 16, "bold"))
    vcc_label.grid(row=0, column=0, padx=(40, 10), pady=(40, 10))
    
    vcc_min_entry = ctk.CTkEntry(new_window, width=100)
    vcc_min_entry.grid(row=0, column=1, padx=(10, 40), pady=(40, 10))
    vcc_min_entry.insert(0, store[0][1])  # Set min value

    vcc_max_entry = ctk.CTkEntry(new_window, width=100)
    vcc_max_entry.grid(row=0, column=2, padx=(10, 40), pady=(40, 10))
    vcc_max_entry.insert(0, store[0][0])  # Set max value

    led_voltage_label = ctk.CTkLabel(new_window, text="LED Voltage(min,max)", font=("Arial", 16, "bold"))
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
    def __init__(self, parent, size=200, fg_color="#4CAF50", bg_color="#E0E0E0", **kwargs):
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

        self.theme_var = ctk.StringVar(value="light")
        ctk.set_appearance_mode(self.theme_var.get())
        ctk.set_default_color_theme("blue")

        self.center_window(self.root, 700, 500)

        self.master_frame = ctk.CTkFrame(master=self.root)
        self.master_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.header_frame = ctk.CTkFrame(master=self.master_frame, height=50)
        self.header_frame.pack(fill="x", pady=(0, 20))

        self.main_label = ctk.CTkLabel(master=self.header_frame, text="USB HUB ZIG", font=("Arial", 24, "bold"))
        self.main_label.pack(side="left", padx=20)

        self.theme_switch = ctk.CTkSwitch(master=self.header_frame, text="Dark Mode", command=self.toggle_theme,variable=self.theme_var, onvalue="dark", offvalue="light")
        self.theme_switch.pack(side="right", padx=20)

        self.content_frame = ctk.CTkFrame(master=self.master_frame)
        self.content_frame.pack(expand=True, fill="both")

        self.widget_frame = ctk.CTkFrame(master=self.content_frame, width=300, height=200, corner_radius=10)
        self.widget_frame.pack(side="left", padx=(0, 20), fill="y")

        self.admin_call = ctk.CTkButton(master=self.widget_frame, text="Admin Login", font=("Arial", 16), corner_radius=8,
                                      height=30, command=lambda: open_win())
        self.admin_call.pack(pady=60)

        self.start_button = ctk.CTkButton(master=self.content_frame, text="Start Test", command=lambda: [self.toggle_status()],
                                       font=("Arial", 16), height=40, corner_radius=8)
        self.start_button.pack(pady=(0, 20))

        self.create_table()
        self.status_widget = CircularStatus(parent=self.content_frame, size=200)
        self.status_widget.pack(side="left", padx=(0, 20), pady=(0, 20))  # Position status widget left

        self.progress_bar = ctk.CTkProgressBar(master=self.content_frame,determinate_speed=0.1,width=400,height=20,corner_radius=0,fg_color="white")
        self.progress_bar.place(x=680,y=600)
        self.progress_bar.set(0)
        self.identifier = None
        self.count = None

        self.create_port_monitor() 

        # Create the COM port monitor
        self.check_devices()  # Start checking for devices


            # If no new ports have been detected for 'stability_time' seconds, or if we've waited for 'max_wait_time', return the ports
    def create_port_monitor(self):
        self.port_monitor_frame = ctk.CTkFrame(master=self.content_frame, corner_radius=10)
        self.port_monitor_frame.pack(side="left", fill="y", padx=(20, 0), pady=(20, 0))

    # Frame for label and GIF
        self.label_gif_frame = ctk.CTkFrame(master=self.port_monitor_frame)
        self.label_gif_frame.pack(pady=(10, 10))

    # Label for active COM ports
        self.port_label = ctk.CTkLabel(master=self.label_gif_frame, text="Active COM Ports", font=("Arial", 16, "bold"))
        self.port_label.pack(side="left", padx=(0, 10))
    # Textbox for displaying active COM ports
        self.port_listbox = ctk.CTkTextbox(master=self.port_monitor_frame, height=200, state="disabled")
        self.port_listbox.pack(expand=True, fill="both")

        self.scanning_label = ctk.CTkLabel(master=self.port_monitor_frame, text="Scanning...", font=("Arial", 14))
        self.scanning_label.pack(pady=(10, 10))

            
    def update_port_list(self):
        

        self.port_listbox.configure(state="normal")  # Enable editing to update
        self.port_listbox.delete("1.0", "end")  # Clear the current list
        com_ports = list_active_ft232r_devices()  # Get active COM ports
        if not com_ports:
            self.port_listbox.insert("end", "No active COM ports.")
        else:
            
            if  self.count == 99:
                self.progress_bar.set(100)
            else:
                self.progress_bar.start()
            self.disable_button()
            for port in com_ports:
                
                self.port_listbox.insert("end", f"{port}\n")
                if len(com_ports) < 3:
                    if self.count != 99:
                        self.count = int(self.progress_bar.get()*100)
                        self.scanning_label.configure(text="Scanning...")
                else:
                    self.progress_bar.stop()
                    self.progress_bar.set(100)
                    self.scanning_label.configure(text="")
                    self.enable_button()
                    

        self.port_listbox.configure(state="disabled") 

    def check_devices(self):
        com_ports = list_active_ft232r_devices()
        if not com_ports:
            self.clear_table()  # Clear the table if no devices are connected
        self.update_port_list()  # Update the port list every check
        
        self.identifier=self.root.after(100, self.check_devices)  # Check every 3 seconds

        print(self.count)
        if self.count == 99:
            print('rUNING')

            root.after_cancel(self.identifier)
            self.progress_bar.stop()
            self.progress_bar.set(100)
            self.enable_button()
            self.root.after(100, self.check_devices)

    
    def show_loading(self, message):
        self.loading_label = ctk.CTkLabel(self.content_frame, text=message)
        self.loading_label.pack()
        self.root.update_idletasks()

    def hide_loading(self):
        if hasattr(self, 'loading_label'):
            self.loading_label.destroy()
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
        data = [["Test Details", "USB Port", "Status","Max Value","Min Value"]]
        for i, port in enumerate(com_ports):
            data.append(["", "", "","N/A","N/A"])

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
        com_ports = list_active_ft232r_devices()
        self.hide_loading()
        if not com_ports:
            self.clear_table()  # Clear the table if no COM ports are available
            messagebox.showinfo("No Ports", "No COM ports detected. Please check the connection.")
            self.enable_button()  # Re-enable the button
            return
        data = [["Test Description", "USB Port/Voltage", "Status","Max Value","Min Value"]]
        DEL=["USB1","USB2","USB3","USB4"]
        
        vcc_check_done=False
        led_check_done=False
        test_results = []
        for i, port in enumerate(com_ports):
            try:
                ser = pyserial.Serial(port, 9600, timeout=2, write_timeout=2)
                time.sleep(0.1)
                ser.reset_input_buffer()
                ser.reset_output_buffer()
                if send_receive_data(ser, packets()[0], active_device=port):
                    val = identifier()
                    if val == False:
                        test_results.append(["-", "-", "No Device Found","N/A","N/A"])
                        self.status_widget.animate("FAIL")        
                    else:
                        volt=voltage_check(ser,packets()[1],active_device=port)
                        main=open("settings.csv","r")
                        store=[]
                        read=csv.reader(main)
                        for i in read:
                            if i==[]:
                                pass
                            else:
                                store.append(i)

                        if not led_check_done:
                            led_volt=voltage_check(ser,packets()[2],active_device=port)
                            if float(store[1][1])>=float(led_volt) and float(store[1][0])<=float(led_volt):
                                test_results.append(["LED Voltage",led_volt,"PASS",store[1][1],store[1][0]])
                            else:
                                test_results.append(["LED Check",led_volt,"FAIL",store[1][0],store[1][1]])
                                self.status_widget.animate("FAIL")
                        led_check_done=True

                        if not vcc_check_done:
                            if float(store[0][1])<=float(volt) and float(store[0][0])>=float(volt):
                                test_results.append(["VCC Check",volt,"PASS",store[0][1],store[0][0]])
                            else:
                                test_results.append(["VCC Check",volt,"FAIL",store[0][1],store[0][0]])
                                self.status_widget.animate("FAIL")
                        vcc_check_done=True           

                    x="USB"+USB_number()
                    DEL.remove(x)
                    test_results.append([f"Test for {x}", val, "PASS","N/A","N/A"])
                else:
                    test_results.append(["-", "-", "FAIL","N/A","N/A"])
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
            result = next((r for r in test_results if test_type in r[0]), [f"Test for {test_type}", "-", "Not Tested","N/A","N/A"])
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
                    self.progress_bar.set(0)

    def toggle_theme(self):
        ctk.set_appearance_mode(self.theme_var.get())





if __name__ == "__main__":

    root = ctk.CTk()
    app = Ziggui(root)
    root.mainloop()