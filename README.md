# USBResetService
A service that resets USB devices based on their location info.
Codes are based on python and powershell commands.

** Open CMD with admin authority and follow the steps below.

1- First, you need to install the pywin32 package:
pip install pywin32

2- Run the command to install the service:
python usb_reset_service.py install                  # Install the service and set "Startup Type: Manuel"

or

python usb_reset_service.py install --startup auto   # Install the service and set "Startup Type: Automatic"

3- To start the service:
python usb_reset_service.py start

4- To stop the service:
python usb_reset_service.py stop

5- To remove the service:
python usb_reset_service.py remove

or

sc delete usbresetservice
