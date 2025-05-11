
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class TextLogger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')
        self.text_widget.update_idletasks()

    def flush(self):
        pass

def create_project_and_devices(project_path, project_name, excel_file, version):
    try:
        import clr
        dll_path = os.path.join(
            "C:\\Program Files\\Siemens\\Automation\\Portal V" + version,
            "PublicAPI",
            "V" + version,
            "Siemens.Engineering.dll"
        )
        clr.AddReference(dll_path)

        import Siemens.Engineering as tia
        import Siemens.Engineering.HW.Features as hwf
        from System.IO import DirectoryInfo
        from Siemens.Engineering import IEngineeringServiceProvider

        print("🚀 Starting TIA Portal with UI...")
        tia_app = tia.TiaPortal(tia.TiaPortalMode.WithUserInterface)

        project_dir = DirectoryInfo(project_path)
        print(f"\n📁 Creating new TIA project: {project_name} at {project_path}")
        project = tia_app.Projects.Create(project_dir, project_name)

        df = pd.read_excel(excel_file)

        module_pairs = []
        for col in df.columns:
            if col.lower().startswith("moduleordernumber"):
                name_col = col.replace("OrderNumber", "Name")
                module_pairs.append((col, name_col if name_col in df.columns else None))

        network_interfaces = []

        for idx, row in df.iterrows():
            device_type = str(row['DeviceType']).strip()
            device_name = str(row['DeviceName']).strip()
            mlfb = "OrderNumber:" + str(row['MLFB']).strip()
            ip = str(row['IP']).strip()
            subnet_name = str(row['SubnetName']).strip()

            print(f"\n📦 Creating device: {device_type} - {device_name}")
            device = project.Devices.CreateWithItem(mlfb, device_name, None)

            for order_col, name_col in module_pairs:
                order_val = row.get(order_col)
                if pd.isna(order_val):
                    continue
                module_order = "OrderNumber:" + str(order_val).strip()
                module_name = str(row[name_col]).strip() if name_col and not pd.isna(row[name_col]) else ""

                plugged = False
                for parent_item in device.DeviceItems:
                    if plugged:
                        break
                    for interface in ["IO1", "PM", "P1"]:
                        if plugged:
                            break
                        for pos in range(1, 15):
                            try:
                                if parent_item.CanPlugNew(module_order, interface, pos):
                                    new_item = parent_item.PlugNew(module_order, interface, pos)
                                    if module_name:
                                        new_item.Name = module_name
                                    print(f"🔌 Plugged {module_order} into {interface} at position {pos}")
                                    plugged = True
                                    break
                            except:
                                continue

                if not plugged:
                    print(f"❌ Could not plug module {module_order}")

            for item in device.DeviceItems:
                for subitem in item.DeviceItems:
                    try:
                        net_service = IEngineeringServiceProvider(subitem).GetService[hwf.NetworkInterface]()
                        if isinstance(net_service, hwf.NetworkInterface):
                            network_interfaces.append(net_service)
                    except:
                        continue

        for i, iface in enumerate(network_interfaces):
            if i >= len(df):
                break
            try:
                iface.Nodes[0].SetAttribute('Address', df.iloc[i]['IP'])
                if i == 0:
                    subnet = iface.Nodes[0].CreateAndConnectToSubnet(df.iloc[i]['SubnetName'])
                    iosystem = iface.IoControllers[0].CreateIoSystem("PNIO")
                else:
                    iface.Nodes[0].ConnectToSubnet(subnet)
                    if iface.IoConnectors.Count > 0:
                        iface.IoConnectors[0].ConnectToIoSystem(iosystem)
            except Exception as net_err:
                print(f"⚠️  Network config error: {net_err}")

        print("\n✅ Project and devices created successfully!")
        messagebox.showinfo("Success", "Project and devices created successfully!")

    except Exception as e:
        print(f"\n❌ Error: {e}")
        messagebox.showerror("Error", str(e))

# ----------------- GUI -----------------
root = tk.Tk()
root.title("TIA Project Creator")
root.geometry("960x720")
root.configure(bg="#1e1e1e")

style = ttk.Style(root)
style.theme_use("clam")
style.configure("TFrame", background="#1e1e1e")
style.configure("TLabel", background="#1e1e1e", foreground="#ffffff", font=("Segoe UI", 11))
style.configure("TEntry", fieldbackground="#2d2d2d", foreground="#ffffff", font=("Segoe UI", 11))
style.configure("TButton", background="#3c3f41", foreground="#ffffff", font=("Segoe UI", 11), padding=8)
style.map("TButton", background=[("active", "#5c5f61")])

project_path_var = tk.StringVar()
project_name_var = tk.StringVar()
excel_file_var = tk.StringVar()
tia_version_var = tk.StringVar(value="18")  # Default to V18

frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

for i in range(3):
    frame.columnconfigure(i, weight=1)

def browse_project_path():
    project_path_var.set(filedialog.askdirectory())

def browse_excel_file():
    excel_file_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")]))

def start_creation():
    create_project_and_devices(
        project_path_var.get(),
        project_name_var.get(),
        excel_file_var.get(),
        tia_version_var.get()
    )

# Input Fields
ttk.Label(frame, text="Project Path:").grid(row=0, column=0, sticky='e', pady=8, padx=5)
ttk.Entry(frame, textvariable=project_path_var, width=70).grid(row=0, column=1, pady=8, sticky='ew')
ttk.Button(frame, text="Browse", command=browse_project_path).grid(row=0, column=2, padx=5)

ttk.Label(frame, text="Project Name:").grid(row=1, column=0, sticky='e', pady=8, padx=5)
ttk.Entry(frame, textvariable=project_name_var, width=70).grid(row=1, column=1, pady=8, sticky='ew')

ttk.Label(frame, text="Excel File:").grid(row=2, column=0, sticky='e', pady=8, padx=5)
ttk.Entry(frame, textvariable=excel_file_var, width=70).grid(row=2, column=1, pady=8, sticky='ew')
ttk.Button(frame, text="Browse", command=browse_excel_file).grid(row=2, column=2, padx=5)

ttk.Label(frame, text="TIA Portal Version:").grid(row=3, column=0, sticky='e', pady=8, padx=5)
ttk.Entry(frame, textvariable=tia_version_var, width=70).grid(row=3, column=1, pady=8, sticky='ew')
ttk.Label(frame, text="e.g., 17, 18, 19").grid(row=3, column=2, padx=5)

ttk.Button(frame, text="🚀 Create Project", command=start_creation).grid(row=4, column=1, pady=20)

# Log Output
ttk.Label(frame, text="Log Output:").grid(row=5, column=0, columnspan=3, sticky='w', pady=(10, 0))

log_frame = ttk.Frame(frame)
log_frame.grid(row=6, column=0, columnspan=3, sticky='nsew', pady=5)
frame.rowconfigure(6, weight=1)

log_text = tk.Text(log_frame, wrap='word', height=20, bg="#2d2d2d", fg="#dcdcdc", insertbackground="white", font=("Consolas", 11))
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
log_text.configure(state='disabled')

scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.configure(yscrollcommand=scrollbar.set)

sys.stdout = TextLogger(log_text)
sys.stderr = TextLogger(log_text)

root.mainloop()
