# TIA Project Creator

The **TIA Project Creator** is a Python application that allows users to create TIA Portal projects and import devices and hardware modules based on data from an Excel file. The tool automates the process of setting up devices, configuring network interfaces, and plugging modules into devices within a TIA Portal project.

## Features

- Create TIA Portal projects with a user interface (UI).
- Import devices and hardware modules from an Excel file.
- Automatically assign IP addresses and subnet information to devices.
- Plug modules into devices and configure their settings based on Excel input.
- Log output and error messages in real-time.

## Requirements

To run the TIA Project Creator, you will need:
- **Python 3.x** installed.
- **TIA Portal** installed on your machine.
- **Python dependencies** listed in `requirements.txt`.

### Dependencies
- pandas
- pythonnet

### Install Dependencies

Before running the script, install the required dependencies using `pip`:

```bash
pip install -r requirements.txt
```

## Setup

1. Clone or Download the Repository:

- Clone this repository or download the zip file containing the project.

2. Prepare the Template Excel File:

- The project requires an Excel file (template.xlsx) with specific columns:

A. DeviceType: Type of the device (e.g., PLC, IO device).

B. DeviceName: Name of the device.

C. MLFB: The order number for the device.

D. IP: The IP address to assign to the device.

E. SubnetName: The name of the subnet to which the device will belong.

F. ModuleOrderNumberX: The order numbers of the modules to plug into the device.

G. ModuleNameX: The names of the modules.

- Example of template.xlsx is in folder

3. Configure TIA Portal Version:

Set the version in the TIA Portal Version input field when using the GUI (defaults to 18).

## Running the Script:

### As a Python Script:

1. Open a terminal or command prompt.

2. Navigate to the directory where the script is located.

3. Run the script:

```bash
python TIA-Project-Creator.py
```

### As a Standalone Executable:

If you're using the compiled executable (TIA-Project-Creator.exe), simply double-click the .exe file to launch the application.

## Use the GUI:

- Project Path: Browse and select the directory where the TIA Portal project will be created.

- Project Name: Enter the name of the project.

- Excel File: Browse and select the Excel file containing the device and module data.

- TIA Portal Version: Select the version of TIA Portal you're using (17, 18, or 19).

- Log Output: View the real-time log output at the bottom of the application.

- Click on Create Project to start the process.

##How it Works

1. Create TIA Project: The script will create a new TIA Portal project in the specified location.

2. Import Devices: Devices are created based on data from the provided Excel file.

3. Assign IPs and Subnets: The IP addresses and subnet names are configured for each device.

4. Plug Modules: Modules are plugged into devices automatically based on the order numbers in the Excel file.

5. Configure Network: Network interfaces and connections are configured based on the device information.

## Example

- Example workflow:

1. Launch the app.

2. Select the project path and give your project a name.

3. Select your template.xlsx file with the required device and module data.

4. Click "Create Project".

5. The script will create the project, import devices, configure network settings, and plug the modules into devices.

6. Watch the log output for the progress and any errors.

## Troubleshooting

1. Error: "TIA Portal not found."
- Ensure that TIA Portal is installed and the correct version is selected in the GUI.

2. Error: "Module cannot be plugged."
- Check if the module order number is valid and if there is enough space in the device to plug the module.

3. Error: "Network configuration error."
- Verify the IP and subnet information in the Excel file.