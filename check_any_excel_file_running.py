import psutil

def is_excel_running():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            # Check if the process is Excel (in Windows, it's usually 'EXCEL.EXE')
            if 'EXCEL.EXE' in proc.info['name'].upper():
                print(f"Excel is running with PID: {proc.info['pid']}")
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

if __name__ == "__main__":
    if is_excel_running():
        print("Excel is currently running.")
    else:
        print("Excel is not running.")
