import tkinter as tk
import os # for basic GUI elements
from tkinter import ttk # for themed widgets
from tkinter import filedialog, messagebox # keep messagebox for errors

# --- Import converter class ---
from pptToPdf import PPTXtoPDFConverter # IMPORTANT: pptToPdf.py should be in the same directory

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX to PDF Converter")
        # Make the window not resizable for simplicity
        self.root.resizable(False, False)
        # Register the closing protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # --- Moved back here: Create the converter instance ---
         # --- Step 2: Create an instance of the converter ---
        try:
            # Instantiate the converter logic. This might take some time if the library is large.
            # If it needs to start PowerPoint. Consider adding a status message later.
            self.converter = PPTXtoPDFConverter()
        except Exception as e:
            # Handle the exception here. You can show an error message to the user.
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.root.destroy() # Close the application
            return
        
        self.create_widgets() # Call the method to create widgets
        
        
    def on_closing(self):
        """Handles the window closing event."""
        print("Close button clicked. Attempting to close converter...") # Debug message
        try:
            if hasattr(self, 'converter') and self.converter:
                self.converter.close() # Call the new close method
        except Exception as e:
            # Log or show error id closing the converter fails
            print(f"Error during converter close: {e}")
            messagebox.showwarning("Cleanup Warning", f"Could not fully close resources:\n{e}")
        finally:
            # Ensure the GUI window is destroyed
            print("Destroying root window.")
            self.root.destroy() # Close the application
        
       
    def create_widgets(self):
        # Create a main frame to hold other widgets
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Configure the main frame's column to expand if window is resized
        main_frame.columnconfigure(0, weight=1)
        
        # --- Widgets ---
        # 1. Single File Button
        self.single_file_button = ttk.Button(
            main_frame,
            text="Convert Single PPTX File",
            command=self.select_and_convert_single # Call the function when button is clicked
        )
        # Place button in row 0, column 0. sticky=(tk.W, tk.E) makes it stretch horizontally
        self.single_file_button.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E))
        
        # 2. Batch Convert Button
        self.batch_convert_button = ttk.Button(
            main_frame,
            text="Convert All PPTX in Folder",
            command=self.select_and_convert_batch # Call the function when button is clicked
        )
        # Place button in row 1, column 0
        self.batch_convert_button.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # 3. Status Label
        self.status_var = tk.StringVar() # Use a StringVar to easily update the label text
        self.status_var.set("Ready. Select an action.") # Initial message
        status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var, # Link the label to the StringVar
            wraplength=350 # Wrap text if it gets too long
        )
        # Place label in row 2, column 0. pady adds space above it.
        status_label.grid(row=2, column=0, padx=5,pady=10, sticky=(tk.W, tk.E))
        
        # --- Add padding to all widgets in the frame ---
        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
    # --- Placeholder functions (we will implement these next) ---
    def select_and_convert_single(self):
        """Handles the conversion of a single PPTX file."""
        self.status_var.set("Select the PowerPoint file to convert.")
        # Force update of the GUI to show the message immediately
        self.root.update_idletasks()
        
        input_path = filedialog.askopenfilename(
            title="Select PowerPoint File",
            filetypes=(("PowerPoint files", "*.pptx *.ppt" ), ("All files", "*.*"))
        )
        
        if not input_path: # User cancelled file selection
            self.status_var.set("Operation cancelled.")
            return
        
        self.status_var.set("Select where to save the PDF...")
        self.root.update_idletasks()
        
        # Suggest a default filename based on the input
        default_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        
        output_path = filedialog.asksaveasfilename(
            title="Save PDF As",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        
        if not output_path: # User cancelled file selection
            self.status_var.set("Operation cancelled. Ready.") # Updated message
            return
        
        # --- Perform Conversion ---
        try:
            self.status_var.set(f"Converting {os.path.basename(input_path)}...")
            #Disable the buttons while conversion is in progress
            self.single_file_button.config(state=tk.DISABLED)
            # Also disable the batch convert button
            if hasattr(self, 'batch_convert_button'):
                self.batch_convert_button.config(state=tk.DISABLED)
            self.root.update_idletasks() # Show status update immediately
            
            success = self.converter.convert_single_file(input_path, output_path, overwrite=True) # Assuming overwrite is okay

            if success:
                self.status_var.set(f"Success! PDF saved to: {output_path}")
                # Ask to open output folder
                if messagebox.askyesno("Success", f"Conversion successful!\nPDF saved to: \n{output_path}\nOpen the output folder?"):
                    try:
                        # Use os.startfile on Windows, or subprocess.call for cross-platform
                        os.startfile(os.path.dirname(output_path))
                    except Exception as e:
                        self.status_var.set(f"Success! PDF saved, but could not open folder: {e}")
                        messagebox.showwarning("Warning", f"Could not open  the output folder: \n{e}")
                        
            else:
                # Check the log file for specific errors if conversion method returns False
                self.status_var.set("Conversion failed. Check conversion.log for details.")
                messagebox.showerror("Conversion Failed", "Could not convert the file. Please check conversion.log for more information.")
        except Exception as e:
            # Catch unexpected errors during conversion process
            self.status_var.set(f"Error during conversion: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
            # Log the error if possible (converter instance might handle logging)
            if hasattr(self.converter, 'logger'):
                self.converter.logger.error(f"GUI Error during single conversion: {e}", exc_info=True)
        
        finally:
            # Re-enable buttons regardless of success or failure
             self.single_file_button.config(state=tk.NORMAL)
             if hasattr(self, 'batch_convert_button'):
                 self.batch_convert_button.config(state=tk.NORMAL)
            # Set status back to ready after completion or error
             if 'Success!' not in self.status_var.get():
                 self.status_var.set("Ready. Select an action")
                
    
    def select_and_convert_batch(self):
        """Handles the conversion of all PPTX files in a folder."""
        self.status_var.set("Select the FOLDER containing  PowerPoint files...")
        self.root.update_idletasks()
        
        input_folder = filedialog.askdirectory(
            title="Select Input Folder with PPTX Files"
        )
        
        if not input_folder: # User cancelled folder selection
            self.status_var.set("Operation cancelled.")
            return
        
        self.status_var.set("Select the FOLDER to save the PDF files...")
        self.root.update_idletasks()
        
        output_folder = filedialog.askdirectory(
            title="Select Output Folder for PDFs"
        )
        
        if not output_folder: # User cancelled folder selection
            self.status_var.set("Operation cancelled.")
            return
        
        # --- Perform Batch Conversion ---
        try:
            self.status_var.set(f"Starting batch conversion in {os.path.basename(input_folder)}...")
            # Disable buttons during conversion
            self.single_file_button.config(state=tk.DISABLED)
            self.batch_convert_button.config(state=tk.DISABLED)
            self.root.update_idletasks() # Show status update immediately
            
            # pass GUI mode to batch_convert method
            # Call the batch convert method from the converter instance
            success = self.converter.batch_convert(input_folder, output_folder, overwrite=True, gui_mode=True) # Assuming overwrite is okay
            
            if success:
                # Ideally, batch_convert should return a list of converted files or success messages
                # For testing, using a generic success message
                self.status_var.set(f"Batch conversion completed for folder: {os.path.basename(input_folder)}")
                if messagebox.askyesno("Success", f"Batch conversion complete!\nPDFs saved in:\n{output_folder}\n\nOpen the output folder?"):
                    try:
                        os.startfile(output_folder)
                    except Exception as e:
                        self.status_var.set(f"Batch conversion complete! Couldn't open folder: {e}")
                        messagebox.showwarning("Warning", f"Could not open the output folder: \n{e}")
            else:
                # If batch_convert returns False, it likely logged the issue.
                self.status_var.set("Batch conversion failed. Check conversion.log for details.")
                messagebox.showerror("Conversion Failed", "Batch conversion failed. Please check conversion.log for more information.")
        except Exception as e:
            # Catch unexpected errors
            self.status_var.set(f"Error during batch conversion: {e}")
            messagebox.showerror("Error", f"An unexpected error occured during batch conversion:\n{e}")
            if hasattr(self.converter, 'logger'):
                self.converter.logger.error(f"GUI Error during batch conversion: {e}", exc_info=True)
        
        finally:
            # Re-enable buttons regardless of success or failure
            self.single_file_button.config(state=tk.NORMAL)
            self.batch_convert_button.config(state=tk.NORMAL)
            # Set status back to ready after completion or error
            if 'Batch conversion' not in self.status_var.get():
                self.status_var.set("Ready. Select an action")
            
                
    def placeholder_action(self):
        print("Button clicked!")
        
    # --- Will add functions for the buttons later ---
    
# Main execution block
if __name__ == "__main__":
    root = tk.Tk() # Create the main window instance
    app = ConverterApp(root) # Create the application instance
    root.mainloop() # Start the Tkinter event loop