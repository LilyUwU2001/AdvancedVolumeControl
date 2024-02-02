# Advanced Volume Control
# Made in the glorious year of 2024 by LilyUwU2001

# Import the necessary TKInter and OS libraries
import tkinter as tk
from tkinter import ttk
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class VolumeControlApp:
    def __init__(self, master):
        self.master = master
        
        # Set window title & sizes
        self.master.title("Advanced Volume Control")
        self.master.geometry("356x232")  # Ustawienie początkowego rozmiaru okna
        self.master.resizable(width=False, height=False)  # Uniemożliwienie zmiany rozmiaru okna
        
        # Add icon to window
        self.master.iconphoto(True, tk.PhotoImage(file="glosnik.png"))

        # Variables to store the checkbox states
        self.checkbox_vars = [tk.IntVar() for _ in range(100)]

        # Create checkboxes, in rows of 10
        self.create_checkboxes()

        # Create the mute checkbox
        self.mute_checkbox = ttk.Checkbutton(self.master, text="Mute", command=self.toggle_mute)
        self.mute_checkbox.grid(row=11, column=0, columnspan=10, sticky="w")
        self.mute_checkbox.state(['!alternate'])  # Make the checkbox start unchecked

    def create_checkboxes(self):
        # Create 100 checkboxes, in rows of 10
        for i in range(1, 101):
            row = (i - 1) // 10
            col = (i - 1) % 10
            checkbox = ttk.Checkbutton(self.master, text=str(i), variable=self.checkbox_vars[i-1], command=self.update_volume)
            checkbox.grid(row=row, column=col, sticky="w")

    def update_volume(self):
        # Calculate the average of the checked checkboxes, round it down, and set the system volume to the result value
        selected_checkboxes = [var.get() for var in self.checkbox_vars]
        selected_count = sum(selected_checkboxes)
        if selected_count > 0:
            sum_volume = sum(i + 1 for i, selected in enumerate(selected_checkboxes) if selected)
            average_volume = sum_volume / selected_count
            volume = round(average_volume / 100, 2)  # Convert to range 0.0 - 1.0
            volume = max(0.0, min(1.0, volume))  # Make sure it's in range 0.0 - 1.0

            # Set the system volume to the result value
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume_object = cast(interface, POINTER(IAudioEndpointVolume))
            volume_object.SetMasterVolumeLevelScalar(volume, None)

    def toggle_mute(self):
        # Mute or unmute the OS depending on the "Mute" checkbox state
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume_object = cast(interface, POINTER(IAudioEndpointVolume))
        if self.mute_checkbox.instate(['selected']):
            # Mute
            volume_object.SetMute(1, None)
        else:
            # Unmute
            volume_object.SetMute(0, None)

if __name__ == "__main__":
    # Run the app and do its main loop
    root = tk.Tk()
    app = VolumeControlApp(root)
    root.mainloop()
