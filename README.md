<h1 align="center" id="title">USB Reset Service</h1>

<p></p>
<p id="description">A service that resets USB devices based on their location info. </p>
<p>Codes are based on python and powershell commands.</p>
<p></p>

<h2>🛠️ Installation Steps:</h2>
<p></p>
<p>1. Open CMD as admin and follow the steps below:</p>

<p>2. First you need to install the pywin32 package:</p>

```
pip install pywin32
```
<p></p>
<p>3. Run the command to install the service:</p>

```
python usb_reset_service.py install                     # Install the service and set "Startup Type: Manuel"

or

python usb_reset_service.py install --startup auto      # Install the service and set "Startup Type: Automatic"
```
<p></p>
<p>4. To start the service:</p>

```
python usb_reset_service.py start
```
<p></p>
<p>5. To stop the service:</p>

```
python usb_reset_service.py stop
```
<p></p>
<p>6. To remove the service:</p>

```
python usb_reset_service.py remove                       # The process may take some time to complete.

or

sc delete usbresetservice                                # Service name 'usbresetservice' is defined in the 'usb_reset_service.py'
```

  
<p> </p>
<p> </p>
<h2>💻 Built with</h2>

Technologies used in the project:

*   Python
*   Powershell
