# 📊 ExcelPoint
**Interactive Roadmap Delivery Confidence Assessment Generator**

ExcelPoint is a Streamlit-based tool that automates the generation of PowerPoint presentations from Excel data. It helps teams assess delivery confidence across multiple releases and instantly produce visually consistent reports with charts, bullet points, and key metrics.

---

## 📌 Features

### 🎯 Core Functionalities
- **Excel-to-PowerPoint Automation**  
  Upload an Excel file with confidence data and a PowerPoint template, and the tool generates a complete presentation.
- **Dynamic Placeholder Replacement**  
  Automatically replaces placeholders (like `<SOLUTION>`, `<Release>`, `<Due Date>`, `<Key>`, `<VALUE>`, `<#>`) with actual data.
- **Chart & Table Updates**  
  Updates charts (High/Medium/Low confidence counts) and text tables directly in the slides.
- **Bullet Point Extraction**  
  Collects confidence-lowering and confidence-increasing factors from team members and fills them into appropriate slides.
- **Multi-Release Support**  
  Handles data for multiple releases in a single Excel sheet.

---

## 🧱 Tech Stack

| Layer        | Technology                |
|--------------|---------------------------|
| Frontend     | Streamlit                 |
| Backend      | Python (pandas, re, python-pptx) |
| File Formats | Excel (.xlsx), PowerPoint (.pptx) |
| Deployment   | Local / Cloud (Streamlit) |
| IDE          | Any Python IDE (VS Code, PyCharm, etc.) |

---

## 🛠️ Setup Instructions

### 🔧 Prerequisites
- Python 3.8+
- pip (Python package manager)
- Excel file with the required data columns
- PowerPoint template with proper placeholders

### 📦 Install Dependencies
```bash pip install streamlit pandas openpyxl python-pptx``` 


## 🚀 Run the Project

1. **Clone the repository**
2. **Run the Streamlit app**
3. **Open the local URL shown in the terminal/ Run streamlit run es.py in cmd**
4. **Upload: The Excel file with release confidence data, The PowerPoint template with placeholders**
5. **Download the Presentation**


---

## 👥 Authors

Ankit Sen — [@senankit501](https://github.com/senankit501)

---

**Made with ❤️ to simplify Excel-based PowerPoint generation**
EOF


   

