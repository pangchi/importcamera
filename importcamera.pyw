title = "Camera File Import and Zip by Year-Month"

import subprocess
import sys
import os
from datetime import datetime
from io import StringIO
import threading
import logging

current_script_path = __file__
current_script_filename = os.path.basename(current_script_path)
current_script_filename_without_extension = os.path.splitext(current_script_filename)[0]

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=f"{current_script_filename_without_extension}.log",
    filemode='w'  # Append mode for log file
)

# Function to install required packages
def install(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        logging.info(f"Successfully installed package: {package}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to install package {package}: {e}")
        raise

packages = ['tkinter', 'zipfile36']
# Ensure required packages are installed
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
    import zipfile
except ImportError:
    try:
        for package in packages:
            install(package)
    except Exception as e:
        logging.error("Failed to install required packages.", exc_info=True)
        print("❌ Error: Failed to install required packages. Please install them manually.")
        sys.exit(1)
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
    import zipfile

# Custom StringIO class to redirect output to ScrolledText and update progress
class TextRedirector(StringIO):
    def __init__(self, text_widget, progress_var=None, total_files=None):
        self.text_widget = text_widget
        self.progress_var = progress_var
        self.total_files = total_files
        self.processed_files = 0

    def write(self, string):
        # Ensure thread-safe Tkinter updates
        self.text_widget.after(0, lambda: self.text_widget.insert(tk.END, string))
        self.text_widget.after(0, lambda: self.text_widget.see(tk.END))  # Auto-scroll
        # Log the output to file
        logging.info(string.strip())
        # Update progress if applicable
        if self.progress_var and "Added " in string and self.total_files:
            self.processed_files += 1
            progress = (self.processed_files / self.total_files) * 100
            self.text_widget.after(0, lambda: self.progress_var.set(progress))

    def flush(self):
        pass  # Required for file-like object compatibility

# ---------------- Core Functionality ----------------
def get_creation_date(file_path):
    try:
        return datetime.fromtimestamp(os.path.getctime(file_path))
    except OSError as e:
        logging.error(f"Failed to get creation date for {file_path}: {e}")
        raise

def zip_files_by_date(directory, outpath, text_widget=None, progress_var=None):
    # Redirect print to text_widget if provided (GUI mode)
    original_stdout = sys.stdout
    total_files = 0
    file_groups = {}

    try:
        # First pass: count total files for progress bar
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            if os.path.isfile(file_path):
                total_files += 1
                try:
                    creation_date = get_creation_date(file_path)
                    folder_name = f"{creation_date.year}-{creation_date.month:02d}"
                    if folder_name not in file_groups:
                        file_groups[folder_name] = []
                    file_groups[folder_name].append(file_path)
                except Exception as e:
                    logging.error(f"Error processing file {file_path}: {e}")
                    print(f"⚠️ Error processing {file_path}: {e}")

        if not file_groups:
            print("⚠️ No files found in input directory.")
            logging.warning(f"No files found in directory: {directory}")
            sys.stdout = original_stdout
            return

        # Set up redirector with progress tracking if in GUI mode
        if text_widget:
            sys.stdout = TextRedirector(text_widget, progress_var, total_files)

        for folder_name, files in file_groups.items():
            zip_filename = os.path.join(outpath, f"{folder_name}.zip")

            counter = 1
            while os.path.exists(zip_filename):
                zip_filename = os.path.join(outpath, f"{folder_name}_{counter}.zip")
                counter += 1

            try:
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file in files:
                        zipf.write(file, os.path.basename(file))
                        print(f"Added {file} to {zip_filename}")
            except Exception as e:
                logging.error(f"Failed to create zip file {zip_filename}: {e}")
                print(f"❌ Error creating {zip_filename}: {e}")

        print("✅ Zipping completed!")
        logging.info("Zipping process completed successfully")

    except Exception as e:
        logging.error("Unexpected error during zipping process", exc_info=True)
        print(f"❌ Unexpected error: {e}")

    finally:
        sys.stdout = original_stdout

# ---------------- Tkinter GUI ----------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(title)

        self.input_dir_var = tk.StringVar(value=r"D:\DCIM")
        self.output_dir_var = tk.StringVar(value=os.getcwd())
        self.progress_var = tk.DoubleVar(value=0)

        tk.Label(root, text="Input Directory:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(root, textvariable=self.input_dir_var, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(root, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(root, text="Output Directory:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(root, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(root, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)

        self.zip_button = tk.Button(root, text="Start Zipping", command=self.run_zip)
        self.zip_button.grid(row=2, column=1, pady=10)

        # Add Progressbar
        self.progress_bar = ttk.Progressbar(root, variable=self.progress_var, maximum=100, length=400)
        self.progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=5)

        # Add ScrolledText widget for console output
        self.output_text = ScrolledText(root, height=10, width=60, wrap=tk.WORD)
        self.output_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    def browse_input(self):
        folder = filedialog.askdirectory(initialdir=self.input_dir_var.get())
        if folder:
            self.input_dir_var.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory(initialdir=self.output_dir_var.get() or os.getcwd())
        if folder:
            self.output_dir_var.set(folder)

    def run_zip(self):
        input_dir = self.input_dir_var.get()
        output_dir = self.output_dir_var.get()

        if not os.path.isdir(input_dir):
            messagebox.showerror("Error", "Invalid input directory.")
            logging.error(f"Invalid input directory: {input_dir}")
            return
        if not os.path.isdir(output_dir):
            messagebox.showerror("Error", "Invalid output directory.")
            logging.error(f"Invalid output directory: {output_dir}")
            return

        # Clear previous output and reset progress
        self.output_text.delete(1.0, tk.END)
        self.progress_var.set(0)
        # Disable the button to prevent multiple clicks
        self.zip_button.config(state="disabled")
        # Run zipping in a separate thread
        threading.Thread(target=self.run_zip_thread, args=(input_dir, output_dir), daemon=True).start()

    def run_zip_thread(self, input_dir, output_dir):
        try:
            # Run the zipping function with progress tracking
            zip_files_by_date(input_dir, output_dir, self.output_text, self.progress_var)
        except Exception as e:
            logging.error("Error in run_zip_thread", exc_info=True)
            self.root.after(0, lambda: messagebox.showerror("Error", f"Zipping failed: {e}"))
        finally:
            # Re-enable the button and show completion message on the main thread
            self.root.after(0, lambda: self.zip_button.config(state="normal"))
            self.root.after(0, lambda: messagebox.showinfo("Done", "Zipping completed!"))

# ---------------- Main Execution ----------------
def running_in_cli():
    try:
        with open('CONIN$', 'r'):
            return True
    except:
        return False

if __name__ == "__main__":
    try:
        if running_in_cli():
            # CLI mode
            if len(sys.argv) != 3:
                print("❌ Error: Invalid number of arguments.")
                print("Usage: python importcamera.pyw <input_directory> <output_directory>")
                logging.error("Invalid number of arguments in CLI mode")
                sys.exit(1)
            else:
                directory_path = sys.argv[1]
                outpath = sys.argv[2]

            if not os.path.isdir(directory_path):
                print("❌ Error: Invalid input directory.")
                logging.error(f"Invalid input directory in CLI mode: {directory_path}")
                sys.exit(1)
            if not os.path.isdir(outpath):
                print("❌ Error: Invalid output directory.")
                logging.error(f"Invalid output directory in CLI mode: {outpath}")
                sys.exit(1)

            zip_files_by_date(directory_path, outpath)
        else:
            # GUI mode
            app = App(tk.Tk())
            app.root.mainloop()
    except Exception as e:
        logging.error("Unexpected error in main execution", exc_info=True)
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)