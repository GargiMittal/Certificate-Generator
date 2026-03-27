# Certificate Generation Script

This script generates personalized certificates using a Word template and data from a CSV file. It replaces placeholders in the template with values from the dataset and combines all generated certificates into a single `.docx` file.

---

## Features

- Bulk certificate generation  
- Customizable `.docx` template  
- Automatic placeholder replacement  
- Single output document  

---

## Requirements

- Python 3.x  
- pandas  
- python-docx  

Install dependencies:

```bash
pip install pandas python-docx
```

---

## Input Files

### CSV (`farewell.csv`)
A CSV file containing the required fields (e.g., Name, Year).

### Template (`test.docx`)

Use placeholders in the document:

_______________________   → Name  
____                      → Year  

---

## Usage

Run the script:

```bash
python script.py
```

Output:

all_certificates.docx

---

## Notes

- Placeholders must match exactly  
- Avoid splitting placeholders across formatted text in Word  
- Works best with simple text templates  
