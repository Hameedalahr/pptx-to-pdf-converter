


# PPT to PDF Conversion in Python  

This repository demonstrates **three different methods** to convert PowerPoint presentations (`.pptx`) into PDF files using Python.  

## Project Structure  

```bash
.
├── PPT_to_PDF_converter-comtypes/
│   ├── venv/
│   ├── requirements.txt
│   └── app.py
├── PPT_to_PDF_converter-pythoncom/  
│   ├── venv/
│   ├── requirements.txt  
│   └── app.py
├── PPT_to_PDF_converter-pythonpptx/
│   ├── venv/
│   ├── uploads/
│   ├── requirements.txt
│   └── app.py
└── README.md
```

## Goal  
Convert `.pptx` → `.pdf` using `pip`-installable libraries while:  
 - The file should not be corrupted during conversion  
 - Use libraries that only can be installed using pip (not brew or out of pip) 
 - Research  about the process by maintaining the requirement.txt file  
 - Should use VSCode and done in virtual environment 

---

## 🔬 Methods Comparison  

| Method               | Cross-Platform | Streamlit-Friendly | Layout Quality | Requires MS PowerPoint | Library Type        |
|----------------------|----------------|--------------------|----------------|------------------------|--------------------|
| `pythoncom` (pywin32) |  Windows only | No               | Excellent    | Yes                  | COM Automation     |
| `comtypes`           |  Windows only | No               | Excellent    | Yes                  | COM Automation     |
| `python-pptx` + `reportlab` |  Yes       | Yes              | May corrupt |  No                   | Pure Python + PDF  |

---

## 1. pythoncom (via pywin32)  
**Pros**:  
- Preserves animations/transitions  
- Perfect layout fidelity  

**Cons**:  
- Windows-only  
- Requires PowerPoint installation
- Streamlit Cloud and most other cloud platforms run on Linux-based containers. Pythoncom rely heavily on Windows COM APIs, which don’t exist on Linux/MacOS.

 [View implementation details](https://github.com/Hameedalahr/pptx-to-pdf-converter/blob/main/PPT%20to%20PDF%20converter%20%20-%20python%20com/readme.md)  

## 2. comtypes  
**Pros**:  
- Same quality as pythoncom  
- Cleaner COM interface  

**Cons**:  
- Same Windows/PowerPoint limitations
- Streamlit Cloud and most other cloud platforms run on Linux-based containers. Pythoncom rely heavily on Windows COM APIs, which don’t exist on Linux/MacOS.

[View implementation details](https://github.com/Hameedalahr/pptx-to-pdf-converter/blob/main/PPT%20to%20PDF%20Converter%20-%20comtypes/readme.md)  

## 3. python-pptx + reportlab  
**Pros**:  
- Works everywhere (Linux/Mac/Windows)  
- Streamlit/Heroku compatible  

**Cons**:  
- Complex layouts may break  
  

 [View implementation details](https://github.com/Hameedalahr/pptx-to-pdf-converter/blob/main/PPT%20to%20PDF%20Converter%20-%20python%20pptx/readme.md)  

---

##  Final Recommendation  

### For Windows users:  
**Best**: `pythoncom` or `comtypes`  
```bash
pip install pywin32
pip install comtypes
```


### For web deployment:  
**Another option**: `python-pptx` + `reportlab`  
```bash
pip install python-pptx reportlab
```

### For local host Running:  
**Run**: `Streamlit` + `Run`  
```bash
streamlit run app.py
```

---

##  Requirements  
- Python 3.8+  
- Virtual environment recommended  
- For COM methods: Microsoft PowerPoint must be installed  

##  Alternatives Considered (But not implemented)  
| Tool          | Reason                          |
|---------------|---------------------------------|
| Aspose        | Commercial license required     |
| CloudConvert  | API limits/authentication needed|

---

##  Summary Table  

| Criterion          | Best Choice              |
|--------------------|--------------------------|
| Layout Quality     | pythoncom/comtypes       |
| Cross-Platform     | python-pptx + reportlab  |
| Free & Pip-Based   | All methods              |
| Cloud Deployment   | Only python-pptx combo   |

