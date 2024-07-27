import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import win32com.client
import shutil

CONFIG_FILE = "config.txt"

class T7PatchLoaderApp:
	def __init__(self, root):
		self.root = root
		self.root.title("")
		self.root.geometry("600x600")
		self.root.configure(bg='#121212')  # Dark background for better readability
		self.center_window()

		self.root.resizable(False, False)  # Disable window resizing

		self.active_window = self.root  # Track the currently active window for dragging

		# Initialize icon presets
		self.icon_presets = {
			'T7+BO3 Icon': 't7bo3.ico',
			'BO3 Icon': 'bo3.ico',
			'T7 Icon': 't7.ico'
		}

		# Create widgets
		self.create_widgets()

		# Initialize the icon presets window tracker
		self.icon_presets_window = None

	def center_window(self):
		"""Center the window on the screen."""
		self.root.update_idletasks()  # Update window size
		width = self.root.winfo_width()
		height = self.root.winfo_height()
		screen_width = self.root.winfo_screenwidth()
		screen_height = self.root.winfo_screenheight()
		x = (screen_width // 2) - (width // 2)
		y = (screen_height // 2) - (height // 2)
		self.root.geometry(f'{width}x{height}+{x}+{y}')

	def create_widgets(self):
		# Create and place the title label at the top of the window
		self.title_label = tk.Label(self.root, text="T7 PATCH LOADER BY SPYRAL", bg='#121212', fg='#ffffff',
								   font=("Refrigerator Deluxe", 24))
		self.title_label.pack(pady=10)  # Adds padding around the title label

		# Create a frame to hold all widgets, positioned below the title label
		self.widget_frame = tk.Frame(self.root, bg='#121212', padx=20, pady=20)
		self.widget_frame.place(relwidth=1, relheight=1, y=65)  # Adjust y to ensure frame is below the title label

		# Action selection
		self.action_label = tk.Label(self.widget_frame, text="SELECT AN ACTION:", bg='#121212', fg='#ffffff', font=("Refrigerator Deluxe", 16, "bold"))
		self.action_label.pack(pady=(0, 0))

		self.action_var = tk.StringVar(value='install')
		self.action_frame = tk.Frame(self.widget_frame, bg='#121212')
		self.action_frame.pack(pady=10)

		self.install_radio = tk.Radiobutton(self.action_frame, text="INSTALL/MODIFY", variable=self.action_var, value='install',
										   bg='#121212', fg='#ffffff', font=("Refrigerator Deluxe", 14), indicatoron=0, 
										   selectcolor='#ff8c33', relief="flat")
		self.install_radio.pack(side='left', padx=20)
		self.uninstall_radio = tk.Radiobutton(self.action_frame, text="UNINSTALL", variable=self.action_var, value='uninstall',
											 bg='#121212', fg='#ffffff', font=("Refrigerator Deluxe", 14), indicatoron=0,
											 selectcolor='#ff8c33', relief="flat")
		self.uninstall_radio.pack(side='left', padx=20)

		# File selection
		self.select_file_button = self.create_custom_button("SELECT T7PATCH_2.03.EXE", self.select_file)

		# Folder creation
		self.shortcut_location_button = self.create_custom_button("SELECT SHORTCUT LOCATION", self.select_shortcut_location)

		# Change Shortcut Name
		self.change_name_button = self.create_custom_button("CHANGE SHORTCUT NAME", self.change_shortcut_name)

		# Change Shortcut Icon
		self.icon_button = self.create_custom_button("CHANGE SHORTCUT ICON", self.show_icon_presets_window)

		# Center-aligned buttons
		self.action_button_frame = tk.Frame(self.widget_frame, bg='#121212')
		self.action_button_frame.pack(pady=40)

		self.install_button = tk.Button(self.action_button_frame, text="EXECUTE", command=self.execute_action,
									   bg='#ff7400', fg='#ffffff', font=("Refrigerator Deluxe", 16, "bold"), width=30)
		self.install_button.pack()

		# Initialize variables
		self.exe_path = ""
		self.folder_path = self.create_folder_in_same_directory()  # Automatically create the folder
		self.shortcut_location = ""
		self.icon_path = self.get_default_icon_path()  # Set a default icon path
		self.shortcut_name = "t7patchloader.lnk"  # Default name

	def create_custom_button(self, text, command):
		"""Create a custom button with hover effects."""
		button = tk.Button(self.widget_frame, text=text, command=command,
						   bg='#ff7400', fg='#ffffff', font=("Refrigerator Deluxe", 14), width=30, height=2)
		button.pack(pady=10)
		button.bind("<Enter>", lambda e: button.config(bg='#ff8c33'))
		button.bind("<Leave>", lambda e: button.config(bg='#ff7400'))
		return button

	def get_default_icon_path(self):
		"""Return the default icon path."""
		script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
		return os.path.join(script_dir, 'icons', 't7bo3.ico')  # Default icon file

	def select_file(self):
		self.exe_path = filedialog.askopenfilename(
			title="Select the path to t7patch_2.03.exe",
			filetypes=[("Executable", "*.exe")]  # Only show .exe files
		)
		if not self.exe_path:
			self.show_warning_message("No file selected!")

	def select_shortcut_location(self):
		self.shortcut_location = filedialog.askdirectory(title="Select Shortcut Location (Defaults to Desktop)")
		if not self.shortcut_location.strip():  # If no directory selected
			self.shortcut_location = os.path.join(os.path.expanduser("~"), "DESKTOP")  # Default to Desktop

	def change_shortcut_name(self):
		new_name = simpledialog.askstring("Change Shortcut Name", "Enter new shortcut name (defaults to t7patchloader):", parent=self.root)
		if new_name:
			self.shortcut_name = new_name + ".lnk"
			self.show_info_message(f"Shortcut name updated to: {self.shortcut_name}")

	def create_main_menu(self):
		"""Create the main menu with a thin orange border."""
		self.main_menu_frame = tk.Frame(self.root, bg='#121212', borderwidth=2, relief='solid')
		self.main_menu_frame.pack(fill='both', expand=True, padx=10, pady=10)

		# Main menu content here
		tk.Label(self.main_menu_frame, text="Main Menu", bg='#121212', fg='#ffffff', font=("Refrigerator Deluxe", 24)).pack(pady=20)
		tk.Button(self.main_menu_frame, text="Open Icon Presets", command=self.show_icon_presets_window, bg='#ff7400', fg='#ffffff').pack(pady=10)

	def show_icon_presets_window(self):
		"""Show a window to select an icon preset."""
		if self.icon_presets_window is None or not tk.Toplevel.winfo_exists(self.icon_presets_window):
			self.icon_presets_window = tk.Toplevel(self.root)
			self.icon_presets_window.title("Select Icon Preset")
			self.icon_presets_window.configure(bg='#121212')
			self.icon_presets_window.resizable(False, False)  # Make the window non-resizable

			# Calculate the position
			window_width = 500
			window_height = 280
			screen_width = self.root.winfo_screenwidth()
			screen_height = self.root.winfo_screenheight()

			# Main window position
			main_window_x = self.root.winfo_x()
			main_window_y = self.root.winfo_y()

			# Centering the icon presets window with an optional offset
			offset_x = 50
			offset_y = 50
			x = main_window_x + (self.root.winfo_width() - window_width) // 2 + offset_x
			y = main_window_y + (self.root.winfo_height() - window_height) // 2 + offset_y

			self.icon_presets_window.geometry(f'{window_width}x{window_height}+{x}+{y}')

			# Create the button frame and add buttons
			button_frame = tk.Frame(self.icon_presets_window, bg='#121212')
			button_frame.pack(pady=15, padx=20, fill='x')

			for preset_name, icon_file in self.icon_presets.items():
				icon_path = self.get_icon_path_from_preset(icon_file)
				icon_image = Image.open(icon_path).resize((32, 32), Image.Resampling.LANCZOS)
				icon_photo = ImageTk.PhotoImage(icon_image)

				button = tk.Button(button_frame, text=preset_name, command=lambda name=preset_name: self.update_icon_from_preset(name),
								   bg='#ff7400', fg='#ffffff', image=icon_photo, compound='left', width=100,
								   font=("Refrigerator Deluxe", 12))
				button.photo = icon_photo  # Keep a reference to avoid garbage collection
				button.pack(pady=10, anchor='center')  # Center the button

			tk.Button(self.icon_presets_window, text="Cancel", command=self.close_icon_presets_window, bg='#ff7400', fg='#ffffff',
					  font=("Refrigerator Deluxe", 12)).pack(pady=20)

	def position_window_relative_to_main(self, window, width, height):
		"""Position the given window relative to the main application window."""
		main_x = self.root.winfo_x()
		main_y = self.root.winfo_y()
		main_width = self.root.winfo_width()
		main_height = self.root.winfo_height()

		# Calculate the new window's position
		x = main_x + (main_width // 2) - (width // 2)
		y = main_y + (main_height // 2) - (height // 2)

		window.geometry(f'{width}x{height}+{x}+{y}')

	def close_icon_presets_window(self):
		"""Close the icon presets window and reset the tracker."""
		if self.icon_presets_window:
			self.icon_presets_window.destroy()
			self.icon_presets_window = None
			# Reset the active window to the main window
			self.active_window = self.root

	def update_icon_from_preset(self, preset_name=None):
		if preset_name is None:
			preset_name = 'T7 Icon'  # Default preset
		icon_file = self.icon_presets.get(preset_name, 't7.ico')
		self.icon_path = self.get_icon_path_from_preset(icon_file)
		self.show_info_message(f"Icon updated to {preset_name}")

	def get_icon_path_from_preset(self, icon_file):
		"""Return the path to the icon file."""
		script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
		return os.path.join(script_dir, 'icons', icon_file)

	def execute_action(self):
		action = self.action_var.get()
		try:
			if action == 'install':
				if not self.exe_path or not self.shortcut_location:
					self.show_warning_message("Please select the executable file and shortcut location!")
					return

				self.create_folder_in_same_directory()  # Ensure the folder is created
				batch_file_path = os.path.join(self.folder_path, 'start_game.bat')
				vbs_file_path = os.path.join(self.folder_path, 'start_game.vbs')
				shortcut_path = os.path.join(self.shortcut_location, self.shortcut_name)

				self.create_batch_file(self.exe_path, batch_file_path)
				self.create_vbs_file(batch_file_path, vbs_file_path)
				self.create_shortcut(vbs_file_path, shortcut_path, self.icon_path)

				installed_items = [self.folder_path, shortcut_path]
				self.log_installed_items(installed_items)
				self.show_info_message("Installation complete!")

			elif action == 'uninstall':
				self.uninstall()

		except Exception as e:
			self.show_error_message(f"An error occurred: {e}")

	def create_batch_file(self, exe_path, batch_file_path):
		with open(batch_file_path, 'w') as batch_file:
			batch_file.write(f'start "" "{exe_path}"\n')
			batch_file.write(f'start "" "steam://rungameid/311210"\n')
			batch_file.write(f'timeout /t 5 /nobreak >nul\n')
			batch_file.write(f':loop\n')
			batch_file.write(f'tasklist /FI "IMAGENAME eq BlackOps3.exe" 2>NUL | find /I /N "BlackOps3.exe">NUL\n')
			batch_file.write(f'if "%ERRORLEVEL%"=="0" (\n')
			batch_file.write(f'    REM BlackOps3.exe IS STILL RUNNING\n')
			batch_file.write(f'    timeout /t 5 /nobreak >nul\n')
			batch_file.write(f'    goto loop\n')
			batch_file.write(f')\n')
			batch_file.write(f'taskkill /IM t7patch_2.03.exe /F\n')
			batch_file.write(f'exit\n')

	def create_vbs_file(self, batch_file_path, vbs_file_path):
		with open(vbs_file_path, 'w') as vbs_file:
			vbs_file.write(f'Set objShell = CreateObject("WScript.Shell")\n')
			vbs_file.write(f'objShell.Run """{batch_file_path}""", 0, False\n')

	def create_shortcut(self, vbs_file_path, shortcut_path, icon_path):
		shell = win32com.client.Dispatch("WScript.Shell")
		shortcut = shell.CreateShortcut(shortcut_path)
		shortcut.TargetPath = vbs_file_path
		shortcut.IconLocation = icon_path
		shortcut.Save()

	def log_installed_items(self, items):
		with open(CONFIG_FILE, 'a') as config_file:
			for item in items:
				config_file.write(f"{item}\n")

	def read_installed_items(self):
		if not os.path.isfile(CONFIG_FILE):
			return []
		with open(CONFIG_FILE, 'r') as config_file:
			return [line.strip() for line in config_file.readlines()]

	def clear_config_file(self):
		open(CONFIG_FILE, 'w').close()  # Clear the file
		os.remove(CONFIG_FILE)  # Remove the config file

	def uninstall(self):
		installed_items = self.read_installed_items()

		if not installed_items:
			self.show_info_message("No installation found to uninstall.")
			return

		for item in installed_items:
			if os.path.isfile(item):
				try:
					os.remove(item)
				except Exception as e:
					self.show_error_message(f"Failed to remove '{item}': {e}")
			elif os.path.isdir(item):
				try:
					shutil.rmtree(item)  # Remove the folder and all its contents
				except Exception as e:
					self.show_error_message(f"Failed to remove folder '{item}': {e}")

		self.clear_config_file()
		self.show_info_message("Uninstallation complete!")

	def create_folder_in_same_directory(self, folder_name="T7PatchLoader"):
		if getattr(sys, 'frozen', False):
			current_directory = os.path.dirname(sys.executable)
		else:
			current_directory = os.path.dirname(os.path.abspath(__file__))

		folder_path = os.path.join(current_directory, folder_name)
		if not os.path.exists(folder_path):
			os.makedirs(folder_path)
		return folder_path

	def show_error_message(self, message):
		"""Display an error message in a popup window."""
		messagebox.showerror("ERROR", message)

	def show_warning_message(self, message):
		"""Display a warning message in a popup window."""
		messagebox.showwarning("Warning", message)

	def show_info_message(self, message):
		"""Display an info message in a popup window."""
		messagebox.showinfo("INFO", message)

if __name__ == "__main__":
	try:
		root = tk.Tk()
		app = T7PatchLoaderApp(root)
		root.mainloop()
	except Exception as e:
		tk.messagebox.showerror("APPLICATION ERROR", f"An unexpected error occurred: {e}")
		with open("error_log.txt", "a") as error_log:
			error_log.write(f"An unexpected error occurred: {e}\n")
