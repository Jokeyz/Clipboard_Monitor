import time
import win32clipboard
import wmi

last_process = None

while True:
    win32clipboard.OpenClipboard()
    clipboard_data = win32clipboard.GetClipboardData()
    clipboard_type = type(clipboard_data)
    win32clipboard.CloseClipboard()

    current_process = None
    c = wmi.WMI()
    for process in c.Win32_Process():
        if process.ProcessId == last_process:
            continue
        try:
            for file in process.ExecutablePath.split('\n'):
                if file.endswith('.exe'):
                    current_process = process
                    break
        except Exception as e:
            pass # ignore processes that we don't have permission to access

    if current_process:
        if current_process != last_process:
            if clipboard_type == str:
                print("Clipboard contents: ", clipboard_data)
            elif clipboard_type == bytes:
                print("Clipboard contents: ", clipboard_data.decode('utf-8'))
            elif clipboard_type == int:
                print("Clipboard contents: ", hex(clipboard_data))
            elif clipboard_type == type(None):
                print("Clipboard is empty")
            else:
                print("Clipboard contents are of type: ", clipboard_type)
            print("Last process to access clipboard: ", current_process.Name)

        # Update the last_process variable
        last_process = current_process.ProcessId

    time.sleep(1) #sleeptime to prevent high cpu usage
