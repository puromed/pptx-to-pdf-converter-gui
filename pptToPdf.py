import os
import logging
from pathlib import Path
import win32com.client
from tqdm import tqdm
from tkinter import filedialog, Tk, messagebox


class PPTXtoPDFConverter:
    def __init__(self, log_file="conversion.log"):
        # setup logging
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        self.logger = logging

        # Initialize Powerpoint application
        try:
            print("Attempting to connect to PowerPoint...")
            self.powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            print("Found running PowerPoint instance")
        except:
            try:
                print("Starting new PowerPoint instance...")
                self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                print("Successfully started PowerPoint")
            except Exception as e:
                self.logger.error(f"Failed to initialize PowerPoint: {str(e)}")
                print("\nTroubleshooting steps:")
                print("1. Verify PowerPoint is installed")
                print("2. Try running the script as administrator")
                print("3. Check if PowerPoint works by opening it manually")
                raise Exception("PowerPoint not installed or not accessible")

    def convert_single_file(self, input_path, output_path=None, overwrite=False):
        try:
            input_path = Path(input_path).resolve()
            if not output_path:
                output_path = input_path.with_suffix(".pdf")
            else:
                output_path = Path(output_path).resolve()
            if not input_path.exists():
                raise FileNotFoundError(f"Input file not found: {input_path}")
            if output_path.exists() and not overwrite:
                self.logger.warning(f"Output file already exists: {output_path}")
                return False

            presentation = self.powerpoint.Presentations.Open(str(input_path))
            presentation.SaveAs(str(output_path), 32)
            presentation.Close()

            self.logger.info(f"Successfully converted: {input_path} -> {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"Error converting {input_path}: {str(e)}")
            return False

    def batch_convert(self, input_folder, output_folder=None, overwrite=False, gui_mode=False):
        """Convert all PPTX files in a folder. Disables tqdm output in GUI mode."""
        try:
            input_folder = Path(input_folder).resolve()
            if output_folder:
                output_folder = Path(output_folder).resolve()
                output_folder.mkdir(parents=True, exist_ok=True)
            else:
                output_folder = input_folder

            # Get all PPTX files
            pptx_files = list(input_folder.glob("*.pptx"))

            if not pptx_files:
                # REMOVED: print(f"No PPTX files found in {input_folder}") 
                self.logger.warning(f"No PPTX files found in {input_folder}")
                return False # Indicate no files found or failure

            # Process files with progress bar
            success_count = 0
            # REMOVED: print(f"\nFound {len(pptx_files)} PPTX files in {input_folder}")

            # Process files with progress bar - disable output in  GUI mode
            # Pass disable=gui_mode to tqdm constructor
            with tqdm(total=len(pptx_files), desc="Converting", disable=gui_mode) as pbar:
                for pptx_file in pptx_files:
                    output_path = output_folder / pptx_file.with_suffix('.pdf').name
                    #REMOVED: print(f"\nConverting: {pptx_file.name}")
                    self.logger.info(f"Batch converting: {pptx_file.name}") # Log instead of print

                    # Use a flag for clarity on conversion success/failure per file
                    file_converted = self.convert_single_file(pptx_file, output_path, overwrite)
                    if file_converted:
                        success_count += 1
                    # Still update pbar even if conversion failed,  to advance progress
                    pbar.update(1)

            #REMOVED:print(f"\nBatch conversion completed: {success_count}/{len(pptx_files)} files converted successfully")
            self.logger.info(
                f"Batch conversion completed. {success_count}/{len(pptx_files)} files converted successfully")
            return True

        except Exception as e:
            error_msg = f"Error in batch conversion: {str(e)}"
            #REMOVED: print(f"\nError: {error_msg}")
            self.logger.error(error_msg)
            return False
        
    def close(self):
        """Safely quits the PowerPoint application instance if it was created."""
        try:
            if self.powerpoint:
                self.powerpoint.Quit()
                print("PowerPoint application closed")
        except Exception as e:
            # Log potential errors during quitting, but don't crash the app
            print(f"Warning: Error while closing PowerPoint: {str(e)}")
            if hasattr(self, 'logger'):
                self.logger.warning(f"Error during PowerPoint Quit: {str(e)}")
                
        finally:
            # Ensure the reference is cleared
            self.powerpoint = None

    # def __del__(self):
    #     try:
    #         if self.powerpoint:
    #             self.powerpoint.Quit()
    #             print("PowerPoint application closed")
    #     except(AttributeError, Exception) as e:
    #         self.logger.warning(f"Error while closing PowerPoint: {str(e)}")


def select_file(title="Select file", file_types=(("PowerPoint files", "*.pptx"),)):
    """Create a file dialog for selecting files"""
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Make dialog appear on top
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=file_types
    )
    root.destroy()
    return file_path


def select_directory(title="Select folder"):
    """Create a dialog for selecting directories"""
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Make dialog appear on top
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path



def main():
    try:
        converter = PPTXtoPDFConverter()

        # Get the current working directory
        current_dir = os.getcwd()
        print(f"\nWorking directory: {current_dir}")

        # Ask user what they want to do
        print("\nWhat would you like to do?")
        print("1. Convert single PPTX file")
        print("2. Batch convert all PPTX files in a folder")
        choice = input("\nEnter your choice (1 or 2): ")

        if choice == "1":
            # Single file conversion with file dialog
            print("\nPlease select the PowerPoint file to convert...")
            input_path = select_file()

            if input_path:  # Check if a file was selected
                # Ask for output location
                print("\nSelect where to save the PDF...")
                suggested_name = Path(input_path).stem + ".pdf"
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    initialfile=suggested_name,
                    filetypes=[("PDF files", "*.pdf")]
                )

                if output_path:  # Check if output location was selected
                    print(f"\nConverting: {input_path}")
                    if converter.convert_single_file(input_path, output_path):
                        print(f"\nSuccess! PDF created at: {output_path}")
                        # Open the output folder
                        os.startfile(os.path.dirname(output_path))
                    else:
                        print("\nConversion failed. Check the log file for details.")
                else:
                    print("\nOperation cancelled by user.")
            else:
                print("\nNo file selected.")

        elif choice == "2":
            # Batch conversion with folder dialog
            print("\nPlease select the folder containing PowerPoint files...")
            input_folder = select_directory(title="Select Input Folder")

            if input_folder:  # Check if input folder was selected
                print("\nPlease select where to save the converted PDFs...")
                output_folder = select_directory(title="Select Output Folder")

                if output_folder:  # Check if output folder was selected
                    print(f"\nInput folder: {input_folder}")
                    print(f"Output folder: {output_folder}")
                    converter.batch_convert(input_folder, output_folder)

                    # Ask if user wants to open output folder
                    if os.path.exists(output_folder):
                        os.startfile(output_folder)
                else:
                    print("\nNo output folder selected. Operation cancelled.")
            else:
                print("\nNo input folder selected. Operation cancelled.")

        else:
            print("\nInvalid choice!")

    except Exception as e:
        print(f"\nError: {str(e)}")
        logging.error(f"Application error: {str(e)}")
        messagebox.showerror("Error", str(e))

    finally:
        print("\nScript execution completed")


if __name__ == "__main__":
    main()

