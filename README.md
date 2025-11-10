# Auto Shalish

<div align="center">  
<img src="https://github.com/user-attachments/assets/57ee86c5-2273-4620-975e-e907540db413" width="600"/>
</div>

A data-driven web application for military HR data that automates the reconciliation, validation, and updating of personnel Excel reports for military HR workflows.  

📹 A short [example video](HaShalish_video.mp4) demonstrates the full process: uploading company and battalion reports, running automated validation and reconciliation, and generating updated Excel files and comment summaries.

---

## Overview  

Traditional HR reporting in large organizations often relies on manual Excel workflows that are time-consuming, error-prone, and lack centralized visibility.  
**Auto Shalish** provides an automated, data-driven solution that consolidates and validates personnel reports across multiple departments or units, ensuring data integrity and operational efficiency.

The system executes:  
- Automatic comparison and synchronization between multiple Excel sources  
- Rule-based validation for personnel and attendance data  
- Anomaly detection and flagging of inconsistent or missing records  
- Generation of color-coded summary reports for fast review and follow-up  

---

## Features  

• **Automated Data Integration:**  
  Automatically merges company-level Excel files into the central battalion report, aligning shared columns, soldier records, and structural consistency across all inputs.  

• **Identity Validation:**  
  Detects duplicate or invalid soldier IDs (non-numeric or incorrect length) and flags any cases requiring manual verification before consolidation.  

• **Missing Data Detection:**  
  Identifies soldiers missing from either report, automatically adds missing records from one file to another, and logs each addition in the comment report.  

• **Department and Unit Validation:**  
  Checks that every soldier’s primary and secondary units match official structure values and flags any unrecognized or invalid department names.  

• **Data Consistency Checks:**  
  Compares overlapping columns such as names, status fields, and hierarchical units to ensure values match between company and battalion levels.  
  When conflicts occur, the system determines which data source is valid and updates accordingly while preserving a full audit trail.  

• **Status and Attendance Validation:**  
  Applies logic-based rules for status fields (e.g., "V", "Released", "Injured", "Absent") and automatically adjusts statuses when conditions are met.  
  For example, detecting new arrivals, release events, or inconsistencies between attendance and location data.  

• **Anomaly Detection:**  
  Highlights irregular reporting cases such as soldiers marked as active but located off-base, mismatched dates, or invalid daily statuses.  
  Each anomaly is color-coded by severity (High, Medium, Low) and documented in a structured comments report.  

• **Automated Comment Report Generation:**  
  Generates a dedicated Excel report summarizing all issues found during processing, including invalid data, structural mismatches, and missing fields.  
  Each record includes soldier identifiers, unit hierarchy, a description of the issue, and priority level for follow-up.  

• **Color-Coded Excel Output:**  
  Produces fully formatted Excel outputs with conditional highlighting and RTL layout, ready for operational use and review.  

• **User-Friendly Web Interface:**  
  The Streamlit interface allows easy uploading of reports, running validations, and downloading updated Excel outputs with no technical background required.  

---

## System Architecture  

| Layer | Description |
|-------|--------------|
| **Frontend** | Streamlit-based interactive interface |
| **Processing** | Python modules handling validation, anomaly detection, and reconciliation |
| **Data Layer** | Pandas and OpenPyXL for structured Excel manipulation |
| **Output** | Auto-generated Excel files with tracked updates and color-coded comment sheets |

---

## How It Looks  

**Start Screen**  
<div align="center">  
<img src="https://github.com/user-attachments/assets/916c35cd-1a06-4cc6-8e43-915da84ced13" width="600"/>
</div>

**Updating Main Report**  
<div align="center">  
<img src="https://github.com/user-attachments/assets/da454158-e1cd-4a0b-abfe-1a53b2fdb6eb" width="600"/>
</div>

**Comments Report**  
<div align="center">  
<img src="https://github.com/user-attachments/assets/adbbbadb-7de2-42c2-8a85-c67c2858f3a8" width="600"/>
</div>

**Comments Excel Output**  
<div align="center">  
<img src="https://github.com/user-attachments/assets/1e0cb18e-ad92-4bdb-a20c-f4a1b7ce92c4" width="600"/>
</div>

---

## 💡 Key Benefits  

- Reduces manual Excel processing time by **over 80%**  
- Prevents human error through structured and rule-based validation  
- Provides full transparency and traceability for every record change  
- Generates a color-coded follow-up report for efficient HR handling  
- Enables faster and more accurate decision-making within the HR chain of command  
- Requires no technical expertise and is ready to use via a simple web interface  

---

## What I Learned  

- Designed an **end-to-end system**, from concept and requirement gathering to full implementation using Python and Streamlit.  
- Acted as both the **developer and the end user**, identifying real operational needs and translating them into practical, data-driven logic.  
- Built and refined the system through iterative testing and real-time feedback during active deployment.  
- Gained experience in **deploying and integrating** a new digital tool in a **non-technical environment**, ensuring usability and adoption.  
- Learned the importance of **clear communication, intuitive UX, and onboarding**, especially when introducing automation tools to field personnel.  
- Observed how, within days, the team **adopted Auto Shalish** as part of their daily workflow, transforming manual HR tracking into a structured, reliable process.  

---

## Running  

Make sure to install all required packages, and then run:  

```bash
streamlit run main.py
```

👥 This project was created by Shira Aronovich.
