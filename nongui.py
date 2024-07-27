import os
import sys
import win32com.client
import shutil

CONFIG_FILE = "config.txt"

def get_file_location(prompt):
	file_path = input(prompt)
	while not os.path.isfile(file_path):
		print("Invalid file path. Please try again.")
		file_path = input(prompt)
	return file_path

def create_folder_in_same_directory(folder_name="T7PatchLoader"):
	# Get the directory of the current script or executable
	if getattr(sys, 'frozen', False):
		# If the script is run as an executable, this is the folder of the executable
		current_directory = os.path.dirname(sys.executable)
	else:
		# If the script is run as a .py file, this is the folder of the script
		current_directory = os.path.dirname(os.path.abspath(__file__))

	# Define the full path for the new folder
	folder_path = os.path.join(current_directory, folder_name)

	try:
		os.makedirs(folder_path, exist_ok=True)
		print(f"Folder '{folder_path}' created successfully.")
	except Exception as e:
		print(f"An error occurred: {e}")
	
	return folder_path

def create_batch_file(exe_path, batch_file_path):
	with open(batch_file_path, 'w') as batch_file:
		batch_file.write(f'@echo off\n')
		batch_file.write(f'start "" "{exe_path}"\n')
		batch_file.write(f'start "" "steam://rungameid/311210"\n')
		batch_file.write(f'timeout /t 5 /nobreak >nul\n')
		batch_file.write(f':loop\n')
		batch_file.write(f'tasklist /FI "IMAGENAME eq BlackOps3.exe" 2>NUL | find /I /N "BlackOps3.exe">NUL\n')
		batch_file.write(f'if "%ERRORLEVEL%"=="0" (\n')
		batch_file.write(f'    REM BlackOps3.exe is still running\n')
		batch_file.write(f'    timeout /t 5 /nobreak >nul\n')
		batch_file.write(f'    goto loop\n')
		batch_file.write(f')\n')
		batch_file.write(f'taskkill /IM t7patch_2.03.exe /F\n')
		batch_file.write(f'exit\n')
	print(f"Batch file created at: {batch_file_path}")

def create_vbs_file(batch_file_path, vbs_file_path):
	with open(vbs_file_path, 'w') as vbs_file:
		vbs_file.write(f'Set objShell = CreateObject("WScript.Shell")\n')
		vbs_file.write(f'objShell.Run """{batch_file_path}""", 0, False\n')
	print(f"VBS file created at: {vbs_file_path}")

def create_shortcut(vbs_file_path, shortcut_path, icon_path):
	shell = win32com.client.Dispatch("WScript.Shell")
	shortcut = shell.CreateShortcut(shortcut_path)
	shortcut.TargetPath = vbs_file_path
	shortcut.IconLocation = icon_path
	shortcut.Save()
	print(f"Shortcut created at: {shortcut_path}")

def select_icon():
	icons = {
		"1": ("T7+Black Ops 3 Icon","t7bo3.ico"),
		"2": ("Black Ops 3 Icon","bo3.ico"),
		"3": ("T7 Icon","t7.ico")
	}

	print("Select an icon from the following options:")
	for key, (name, _) in icons.items():
		print(f"{key}: {name}")

	choice = input("Enter the number of your choice: ")
	icon_file = icons.get(choice)

	if not icon_file:
		print("Invalid choice. Using default icon")
		icon_file = icons["1"]  # Default to the first icon

	if isinstance(icon_file, tuple):
		icon_filename = icon_file[1]    
	else:
		print("Unexpected icon file format. Using default icon.")
		icon_filename = icons["1"][1]  # Default to the first icon file

	if not isinstance(icon_filename, str):
		print("Icon file name is not a string. Using default icon.")
		icon_filename = icons["1"][1]

	script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
	icon_path = os.path.join(script_dir, icon_filename)

	if not os.path.isfile(icon_path):
		print(f"Icon file '{icon_path}' not found. Please ensure it exists in the same directory as the executable.")
		sys.exit(1)

	return icon_path

def log_installed_items(items):
	with open(CONFIG_FILE, 'a') as config_file:
		for item in items:
			config_file.write(f"{item}\n")

def read_installed_items():
	if not os.path.isfile(CONFIG_FILE):
		return []
	with open(CONFIG_FILE, 'r') as config_file:
		return [line.strip() for line in config_file.readlines()]

def clear_config_file():
	open(CONFIG_FILE, 'w').close()  # Clear the file
	os.remove(CONFIG_FILE)  # Remove the config file

def uninstall():
	installed_items = read_installed_items()

	if not installed_items:
		print("No installation found to uninstall.")
		return

	for item in installed_items:
		if os.path.isfile(item):
			try:
				os.remove(item)
				print(f"Removed '{item}'.")
			except Exception as e:
				print(f"Failed to remove '{item}': {e}")
		elif os.path.isdir(item):
			try:
				shutil.rmtree(item)  # Remove the folder and all its contents
				print(f"Removed folder '{item}'.")
			except Exception as e:
				print(f"Failed to remove folder '{item}': {e}")

	clear_config_file()

def wait_for_exit():
	input("Operation complete. Press Enter to exit...")

if __name__ == "__main__":
	
	print("Welcome to the T7 Patch Loader")
	action = input("Enter 'install' to install/modify or 'uninstall' to uninstall: ").strip().lower()

	if action == 'install':
		exe_path = get_file_location("Enter the full path to t7patch_2.03.exe: ")
		folder_path = create_folder_in_same_directory()  # No need to prompt for folder name
		
		batch_file_name = "t7patchbooter.bat"
		batch_file_path = os.path.join(folder_path, batch_file_name)
		create_batch_file(exe_path, batch_file_path)

		vbs_file_name = "batloader.vbs"
		vbs_file_path = os.path.join(folder_path, vbs_file_name)
		create_vbs_file(batch_file_path, vbs_file_path)

		shortcut_name = input("Enter the name for the shortcut: ")
		if not shortcut_name.endswith('.lnk'):
			shortcut_name += '.lnk'

		shortcut_location = input("Enter the full path where the shortcut should be placed (leave blank for Desktop): ")
		
		if not shortcut_location.strip():  # If the input is empty or only whitespace
			shortcut_location = os.path.join(os.path.expanduser("~"), "Desktop")  # Default to Desktop

		icon_path = select_icon()
		shortcut_path = os.path.join(shortcut_location, shortcut_name)
		create_shortcut(vbs_file_path, shortcut_path, icon_path)

		# Log installed items
		log_installed_items([folder_path, batch_file_path, vbs_file_path, shortcut_path])

		wait_for_exit()

		wait_for_exit()

	elif action == 'uninstall':
		uninstall()
		wait_for_exit()

	else:
		print("Invalid action. Please enter 'install' or 'uninstall'.")
		wait_for_exit()
