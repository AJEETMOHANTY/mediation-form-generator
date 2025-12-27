# Mediation Application Form Generator

This project generates a professionally formatted **Mediation Application Form (Form A)** in Microsoft Word (`.docx`) format using Python.

The document layout, spacing, and structure closely replicate the provided PDF file, as required in the assignment.


## 📌 Project Overview

The goal of this project is to **recreate a fixed PDF document layout** using Python — not to build a dynamic form system.

All content is intentionally **static**, matching the original document exactly.


## 🧩 Key Features

- Accurate recreation of the provided PDF layout  
- Structured table-based formatting  
- Controlled spacing and alignment  
- Clean and readable Word document  
- Clickable email link ("info@kslegal.co.in")  
- Professional legal-document appearance  


## 🛠️ Technologies Used

- **Python 3.x**
- **python-docx**

*(No external APIs required)*


## 📁 Project Structure
```text
project/
│
├── app.py # Main script to generate the Word document
├── requirements.txt # Python dependencies
├── README.md # Project documentation
│
└── output/
└── Mediation_Form.docx
```


## ▶️ How to Run

### 1️⃣ Install dependencies
pip install -r requirements.txt
### 2️⃣ Run the script
python app.py
### 3️⃣ Output
output/Mediation_Form.docx


## Deployment Note

This project is designed as a document generation utility using Python and `python-docx`.

The application can be deployed on platforms like Koyeb or Railway using Gunicorn.
However, due to platform-specific runtime constraints, deployment may require minor environment configuration.

The core logic, structure, and output generation work correctly when run locally using:

```bash
python app.py
