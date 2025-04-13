# PPTX to PDF Converter GUI

## Project Summary

This project provides a user-friendly Graphical User Interface (GUI) application for Windows users to easily convert Microsoft PowerPoint `.pptx` files into `.pdf` format. It aims to simplify the conversion process compared to manually opening each file in PowerPoint or using command-line tools, especially for converting multiple files at once.

The application leverages Microsoft PowerPoint's own conversion capabilities through automation, ensuring high fidelity in the resulting PDFs.



## Core Features

*   **Single File Conversion:** Select an individual `.pptx` file and choose where to save the converted `.pdf`.
*   **Batch Folder Conversion:** Select an input folder containing multiple `.pptx` files and an output folder to save all the converted `.pdf` files.
*   **Simple Interface:** An intuitive GUI built with Tkinter, requiring no command-line interaction for basic use.
*   **Status Updates:** The application provides feedback on the current operation (e.g., "Selecting file...", "Converting...", "Success!", "Error...").
*   **Logging:** Records details of each conversion attempt (success or failure) into a `conversion.log` file located in the same directory as the executable (or the source script).

## Prerequisites

*   **Operating System:** **Windows** (This tool relies on Windows-specific libraries (`pywin32`) to interact with Microsoft Office).
*   **Microsoft PowerPoint:** A **valid, installed, and activated version** of Microsoft PowerPoint is **required** on the system where you run this converter. The application automates the installed PowerPoint application to perform the actual conversion. It will not work without PowerPoint.

## How to Use the Application (Recommended Method)

This is the easiest way for most users to get started:

1.  **Download the Executable:**
    *   Navigate to the [**Releases**](https://github.com/puromed/pptx-to-pdf-converter-gui/releases/tag/v1.0.0) page of this GitHub repository.
    *   Look for the latest release version (e.g., `v1.0.0`).
    *   Under the "Assets" section for that release, download the file named `PPTXtoPDF_vX.Y.Z.exe` (the version numbers will match the release).

2.  **Run the Converter:**
    *   Locate the downloaded `PPTXtoPDF_vX.Y.Z.exe` file (likely in your Downloads folder).
    *   **Double-click the `.exe` file.** No installation is necessary.
    *   The application window should appear.

3.  **Perform Conversions:**
    *   **For a single file:**
        *   Click the "Convert Single PPTX File" button.
        *   A dialog box will appear; select the `.pptx` file you want to convert.
        *   Another dialog box will appear; choose the location and name for the output `.pdf` file.
        *   The status bar will update during conversion. A message box will confirm success or failure.
    *   **For multiple files in a folder:**
        *   Click the "Convert All PPTX in Folder" button.
        *   A dialog box will appear; select the *folder* containing the `.pptx` files.
        *   Another dialog box will appear; select the *folder* where you want the converted `.pdf` files to be saved.
        *   The status bar will update. A message box will confirm completion or report issues.

4.  **Check Logs (If Needed):** If a conversion fails, check the `conversion.log` file (it will be created in the same location where you ran the `.exe`) for more detailed error messages.

## How the Executable (`.exe`) Works

The provided `.exe` file is created using **PyInstaller**. PyInstaller bundles the Python scripts (`gui_app.py`, `pptToPdf.py`), the necessary Python interpreter components, and required libraries (`pywin32`, `tqdm`, `tkinter`) into a single executable package.

When you run the `.exe`:
*   It unpacks the necessary components (sometimes into a temporary directory).
*   It executes the Python code.
*   Crucially, it uses the `pywin32` library to communicate with your **already installed Microsoft PowerPoint application** via its COM interface. It essentially tells PowerPoint to open the `.pptx` file and save it as a `.pdf`.
*   This means the `.exe` itself does not contain the PowerPoint conversion engine; it *requires* PowerPoint to be present on the system.

## Running from Source (for Developers)

If you want to modify the code or run it directly using Python:

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/puromed/pptx-to-pdf-converter-gui.git
    cd [Your-Repo-Name]
    ```
2.  **Set up Environment (Recommended):** Create and activate a Python virtual environment.
    ```bash
    python -m venv venv
    # On Windows cmd:
    venv\Scripts\activate
    # On Git Bash / Linux / macOS:
    # source venv/bin/activate
    ```
3.  **Install Dependencies:**
    ```bash
    pip install pywin32 tqdm
    ```
    *(Tkinter is typically included with standard Python installations on Windows).*
4.  **Run the Application:**
    ```bash
    python gui_app.py
    ```

## Troubleshooting

*   **"Failed to initialize PowerPoint" Error:** Ensure Microsoft PowerPoint is correctly installed, activated, and can be opened manually. Try running the `.exe` as an administrator (right-click -> Run as administrator) once, although this shouldn't normally be required.
*   **Conversion Fails for Specific Files:** Some complex `.pptx` files might have issues during automated conversion. Check the `conversion.log` for specific error messages from PowerPoint. Try opening and saving the file manually in PowerPoint.
*   **Antivirus Flags:** Occasionally, `.exe` files created by PyInstaller (especially `--onefile` versions) might be flagged by antivirus software. This is often a false positive due to the way PyInstaller bundles applications. If you trust the source (this repository), you may need to create an exception in your antivirus program.


## Contributing

Contributions are welcome! If you'd like to help improve this project, please follow these steps:

  **Report Bugs or Suggest Features:** Use the [GitHub Issues](https://github.com/puromed/pptx-to-pdf-converter-gui/issues) page for this repository to report any bugs you find or suggest new features or improvements. Please provide as much detail as possible.


## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
