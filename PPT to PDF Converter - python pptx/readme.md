



# PPT to PDF Conversion using `python-pptx` and `reportlab`

This method **does not use Microsoft PowerPoint**. It reads PowerPoint slides using `python-pptx` and re-renders them as a PDF using `reportlab`.

> ✅ Works cross-platform but renders only **text content**, not design fidelity.




## 🛠️ How It Works

Parses .pptx file using python-pptx

Extracts slide titles and text content

Uses reportlab to write the content into a PDF

- **Pros**:
  - Best used for automation tasks where visuals are secondary.
  - Cross platform compatabilty.
  - Streamlit deployment suitable

- **Cons**:
  - Not accurate conversion.
  - loss in data
    
## 📦 Requirements

Install `pythonpptx and reportlab` using pip:

```bash
pip install -r requirement.txt

