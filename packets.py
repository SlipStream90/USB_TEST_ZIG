import serial as pyserial
import time
import threading
from com_port import list_active_ft232r_devices
import os
from colorama import Fore, init
import win32com.client
import re
import time


init()

fetching_d4 = {}
fetching_d4_lock = threading.Lock()  # To ensure thread safety


def crc_16bit_modbus(length, buffer):
    crc_temp = 0xFFFF
    for pos in range(length):
        crc_temp ^= buffer[pos]
        for _ in range(8):
            if crc_temp & 1:
                crc_temp >>= 1
                crc_temp ^= 40961
            else:
                crc_temp >>= 1
    return crc_temp


def packets():
    PACKET = []
    board_id = 0xAB
    register_start = 0x01
    get_ = 0x0A
    byte_1 = 0x00
    byte_2 = 0x00

    for register in range(register_start, register_start + 3):
        lt = [board_id, register, get_, byte_1, byte_2]
        check_sum = crc_16bit_modbus(len(lt), lt)
        low_byte = check_sum & 0xFF
        high_byte = (check_sum >> 8) & 0xFF
        lt.extend([low_byte, high_byte])
        PACKET.append(lt)

    for i in range(5):  # Generates packets with the second byte from 0x00 to 0x04
        register = 0x00 + i
        lt = [board_id, register, get_, byte_1, byte_2]
        check_sum = crc_16bit_modbus(len(lt), lt)
        low_byte = check_sum & 0xFF
        high_byte = (check_sum >> 8) & 0xFF
        lt.extend([low_byte, high_byte])
        PACKET.append(lt)    




    return PACKET


def send_receive_data(ser,hex_code):
    print("This is the sent packet: ", hex_code)
    data = bytes(hex_code)

    try:
        ser.write(data)
    except pyserial.SerialTimeoutException:
        print("Write operation timed out")
        os.abort()

    try:
        response = ser.read(7)
    except pyserial.SerialException as e:
        print(f"Error reading from port: {e}")
        return False

    response_hex = [hex(byte)[2:].zfill(2).upper() for byte in response]
    global final_hex
    final_hex = ' '.join(response_hex)

    print("This is the Response packet: ", final_hex)

    if len(response_hex) == 7:
        print(f"Device sent: {final_hex}")
        print("Matched")
        return True
    else:
        print("Can't read data !!")
        return False
    
def voltage_check(ser,hex_data,active_device):
    data=bytes(hex_data)
    try:
        ser.write(data)
    except pyserial.SerialTimeoutException:
        print("Write operation timed out")
        os._exit(1)
    try:
        response = ser.read(7)
    except pyserial.SerialException as e:
        print(f"Error reading from port: {e}")
        os.abort()
    response_hex = [hex(byte)[2:].zfill(2).upper() for byte in response]
    volt_hex = ' '.join(response_hex)
    volt=volt_hex[9]+volt_hex[10]+volt_hex[12]+volt_hex[13]
    volt=int(volt,16)
    volt = float(volt) / 100
    return volt


def com_selector(hex_data):
    com_ports=list_active_ft232r_devices()
    data=bytes(hex_data)
    acc_list=[]
    for i in com_ports:
        try:
            ser = pyserial.Serial(i, 9600, timeout=2)
            ser.write(data)
            response = ser.read(7)
            response_hex = [hex(byte)[2:].zfill(2).upper() for byte in response]
            volt_hex = ' '.join(response_hex)
            if volt_hex==hex_data:
                acc_list.insert(0,i)
            else:
                acc_list.insert(1,i)    


        except Exception as e:
            print(f"Error reading from port: {e}")

def send_data(ser, hex_code):
    print("This is the sent packet: ", hex_code)
    data = bytes(hex_code)

    try:
        ser.write(data) 
        return True    
    except pyserial.SerialTimeoutException:
        print("Write operation timed out")
        os.abort()
   

            






    
def identifier():
    try:
        port_no=final_hex[12]+final_hex[13]
        return port_no
    except IndexError:
        return False

def USB_number():
    try:
        USB_no=final_hex[13]
        return USB_no
    except IndexError:
        return False



def communicate_with_device(port):
    try:
        ser = pyserial.Serial(port, 9600, timeout=2, write_timeout=2)
        time.sleep(0.1)  # Short delay after opening the port
        ser.reset_input_buffer()
        ser.reset_output_buffer()

        if send_receive_data(ser, packets()[0], active_device=port):
            print(f"Successfully communicated with device on port {port}")
        else:
            print(f"Failed to communicate with device on port {port}")

    except pyserial.SerialException as e:
        print(f"Error opening or communicating with port {port}: {e}")

    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print(f"Closed port {port}")

    print("------------------------")


def main():
    active_devices = list_active_ft232r_devices()

    if not active_devices:
        print("No active FT232R devices found.")
        return False

    threads = []

    for port in active_devices:
        print(f"Attempting to communicate with device on port {port}")
        thread = threading.Thread(target=communicate_with_device, args=(port,))
        thread.start()
        threads.append(thread)

    for thread in threads:
        thread.join()

    return True

def list_mice_devices():
    wmi = win32com.client.GetObject("winmgmts:")
    devices = wmi.InstancesOf("Win32_PointingDevice")
    mice_list = []

    for device in devices:
        name = device.Name
        match = re.search(r'VID_(\w+)&PID_(\w+)', device.PNPDeviceID)
        
        vid, pid = match.groups() if match else (None, None)
        mice_list.append((name, vid, pid))
    return mice_list    

def vcc_led_check():
    PACKET = []
    board_id = 0xAB
    register_start = 0x01
    get_ = 0x0A
    byte_1 = 0x00
    byte_2 = 0x00

    for register in range(register_start, register_start + 3):
        lt = [board_id, register, get_, byte_1, byte_2]
        check_sum = crc_16bit_modbus(len(lt), lt)
        low_byte = check_sum & 0xFF
        high_byte = (check_sum >> 8) & 0xFF
        lt.extend([low_byte, high_byte])
        PACKET.append(lt)

    return PACKET    



if __name__ == "__main__":
    x=list_mice_devices()
    usb_port=[tup for tup in x if "USB Input Device" in tup]
    print(len(usb_port))



    