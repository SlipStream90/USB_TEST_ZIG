
import serial.tools.list_ports


def list_active_ft232r_devices(vid=0x0403, pid=0x6001):
    
    """
    List all active FT232R USB UART devices based on the specified VID and PID.

    Args:
        vid (int): Vendor ID of the FT232R USB UART.
        pid (int): Product ID of the FT232R USB UART.

    Returns:
        List[Tuple[str, str]]: A list of tuples where each tuple contains the COM port and device name.
    """

    # Get a list of all available COM ports
    ports = serial.tools.list_ports.comports()
    
    # Filter ports related to the specified VID and PID
    ft232r_ports = [port.device for port in ports if port.vid == vid and port.pid == pid]
    
    return ft232r_ports

#if __name__ == "__main__":
    VENDOR_ID = 0x0403
    PRODUCT_ID = 0x6001

    ft232r_ports = list_active_ft232r_devices(VENDOR_ID, PRODUCT_ID)
    if ft232r_ports:
        print("Active FT232R USB UART COM Ports:")
        for port, description in ft232r_ports:
            print(f"Port: {port}, Device Name: {description}")
    else:
        print("No active FT232R USB UART COM ports found.")



x=list_active_ft232r_devices(vid=0x0403, pid=0x6001)



#while True :
    #print(list_active_ft232r_devices(vid=0x0403, pid=0x6001))



