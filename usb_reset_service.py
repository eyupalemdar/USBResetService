import winreg
import subprocess
import time
import servicemanager
import win32serviceutil
import win32service
import win32event
import logging
import os
from pathlib import Path

class USBResetService(win32serviceutil.ServiceFramework):
    _svc_name_ = "USBResetService"
    _svc_display_name_ = "USB Reset Service"
    _svc_description_ = "A service that resets USB devices based on their location info."
    
    _check_interval_ = 60000 # 60 seconds
    _parent_usb_location_info_ = "Port_#0013.Hub_#0001"
    _lianli_usb_location_info_ = "Port_#0003.Hub_#0004"   # Reset Parent usb if Lian-li usb failed
    
    # Define the log file location
    log_directory = os.path.dirname(os.path.abspath(__file__))
    _log_file_location_ = os.path.join(log_directory, "usb_reset_service.log")

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        
        logging.basicConfig(filename=self._log_file_location_, 
                            level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s")
                            
        # Temporarily disable logging
        logging.disable(logging.CRITICAL) # Turn off this piece of code to enable logging.
                            
        logging.info(f"Service is started")


    def SvcStop(self):
        logging.info("Service is stopped.")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
                              
        try:
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            self.main()
        except Exception as e:
            logging.error(f"Error in SvcDoRun: {str(e)}")
            servicemanager.LogErrorMsg(f"Error in SvcDoRun: {str(e)}")

    def find_usb_device_id_in_registry(self, usb_location_info):
        base_key = r"SYSTEM\CurrentControlSet\Enum\USB"
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_key) as usb_key:
                subkey_count = winreg.QueryInfoKey(usb_key)[0]  

                for i in range(subkey_count):
                    subkey_name = winreg.EnumKey(usb_key, i)
                    with winreg.OpenKey(usb_key, subkey_name) as subkey:
                        subsubkey_count = winreg.QueryInfoKey(subkey)[0]

                        for j in range(subsubkey_count):
                            subsubkey_name = winreg.EnumKey(subkey, j)
                            device_path = f"{base_key}\\{subkey_name}\\{subsubkey_name}"
                            
                            try:
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, device_path) as device_subkey:
                                    location_info = winreg.QueryValueEx(device_subkey, "LocationInformation")[0]
                                    if usb_location_info.lower() == location_info.lower():
                                        return f"USB\\{subkey_name}\\{subsubkey_name}"
                            except FileNotFoundError:
                                continue
            logging.info("DeviceID not found in registry.")
            return None
        except Exception as e:
            logging.error(f"Error accessing registry: {e}")
            servicemanager.LogErrorMsg(f"Error accessing registry: {e}")
            return None

    def reset_usb_device(self, instance_id):
        try:
            disable_command = f'Get-PnpDevice -InstanceId "{instance_id}" | Disable-PnpDevice -Confirm:$false'
            result = subprocess.run(["powershell", "-Command", disable_command], capture_output=True, text=True)
            output = result.stdout.strip()
            if "Error" in output:
                logging.warning(f"Device {instance_id} could not be disabled: {output}")
            else:
                time.sleep(3)

                enable_command = f'Get-PnpDevice -InstanceId "{instance_id}" | Enable-PnpDevice -Confirm:$false'
                result = subprocess.run(["powershell", "-Command", enable_command], capture_output=True, text=True)
                output = result.stdout.strip()
                if "Error" in output:
                    logging.warning(f"Device {instance_id} could not be enabled: {output}")
                else:
                    logging.info(f"Device {instance_id} reset successfully.")
        except subprocess.CalledProcessError as e:
            logging.error(f"subprocess error while resetting device: {e}")
        except Exception as e:
            logging.error(f"Unexpected error in reset_usb_device: {e}")
    
    def is_device_failed(self, instance_id):
        try:
            status_command = f'Get-PnpDevice -InstanceId "{instance_id}"'
            logging.info(f"Powershell command is running: {status_command}")
            result = subprocess.run(["powershell", "-Command", status_command], capture_output=True, text=True)
            output = result.stdout.strip()
    
            if "Error" in output:
                return True
            else:
                return False
        except subprocess.CalledProcessError as e:
            logging.error(f"subprocess error while checking device status: {e}")
            return True
        except Exception as e:
            logging.error(f"Unexpected error in is_device_failed: {e}")
            return True
    
    def main(self):
        while self.running:
            
            try:
                lianli_device_id = self.find_usb_device_id_in_registry(self._lianli_usb_location_info_)
                is_liani_failed = self.is_device_failed(lianli_device_id)
                logging.info(f"lianli_device_id: {lianli_device_id}  failed: {is_liani_failed}  result = {lianli_device_id and is_liani_failed}")
                
                if lianli_device_id and is_liani_failed:
                    parent_device_id = self.find_usb_device_id_in_registry(self._parent_usb_location_info_)
                    if parent_device_id:
                        self.reset_usb_device(parent_device_id)
                    else:
                        logging.warning("Parent device ID could not be found")
                else:
                    logging.warning("Lianli device ID is functioning properly or could not be found")
            except Exception as e:
                logging.error(f"Error in main loop: {e}")
            
            # If it is still working after the number of seconds specified by the _check_interval_ variable, continue checking the devices.
            wait_result = win32event.WaitForSingleObject(self.hWaitStop, self._check_interval_)
            if wait_result == win32event.WAIT_OBJECT_0:
                logging.info("Service is stopping...")
                break

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(USBResetService)
