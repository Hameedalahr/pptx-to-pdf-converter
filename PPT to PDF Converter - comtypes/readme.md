# PPT to PDF Conversion using `comtypes`

This method utilizes the `comtypes` library to automate Microsoft PowerPoint and convert `.ppt` or `.pptx` files into `.pdf` format.

>  This method only works on **Windows** with **Microsoft PowerPoint installed**.

##  How It Works

Uses COM automation to open PowerPoint

Loads the presentation file

Saves it as PDF using PowerPoint's internal API

- **Pros**:
  - Native Office Automation via COM
  - Stable with excellent layout preservation
- **Cons**:
  - Only works on Windows
  - Requires PowerPoint installed
  - Cannot be hosted/deployed via Streamlit or cloud
    
##  Requirements

Install `comtypes and streamlit` using pip:

```bash
pip install -r requirement.txt

