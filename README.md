# JMT-Report-Generator
📊 Janma-Mrityu Tathya Portal – Excel Report Tools

> ⚠️ Disclaimer: These scripts are designed specifically for Excel files exported from the Janma-Mrityu Tathya Portal via  
Report > Application List (Birth) or Report > Application List (Death).

 🔧 Features & Tools Overview

 1. Birth Application Processor – Add Time Gap & Remove Invalid Entries
Processes birth application Excel files by:
- Adding a "Time Gap" column (difference between Date of Birth and Reporting Date) placed between Date of Birth and Gender.
- Filtering out the following:
  - Applications with Status as Application Submitted or Rejected
  - Applications with Appl Type as Still Birth or Old Reg. Event Birth

 2. Death Application Processor – Add Time Gap & Remove Invalid Entries
Processes death application Excel files by:
- Adding a "Time Gap" column (difference between Date of Death and Reporting Date) placed between Date of Birth and Gender.
- Filtering out:
  - Applications with Status as Application Submitted or Rejected

---

 📈 Dashboard Reports

 3. Birth Application Dashboard
Generates interactive visual dashboards using processed birth application data. Key metrics include:
1. Applications by Reporting Date  
2. Applications by Date of Birth  
3. Time Gap Distribution  
4. Missing First Names (Excl. Still Birth)  
5. Place of Birth  
6. Father's ID Proof  
7. Mother's ID Proof  
8. Present Address – District  
9. Present Address – Block/Municipality  
10. Permanent Address – District  
11. Permanent Address – Block/Municipality  
12. Religion  
13. Father's Education  
14. Father's Occupation  
15. Mother's Education  
16. Mother's Occupation  
17. Application Status  
18. Application Type  
19. Type of Attention at Delivery  
20. Delivery Method  
21. Entered By (Role)  
22. Birth Weight Distribution  

 4. Death Application Dashboard
Generates dynamic dashboard reports for death application data. Key insights include:
1. Applications by Reporting Date  
2. Applications by Date of Death  
3. Time Gap Distribution  
4. Age Distribution at Time of Death  
5. Gender Distribution  
6. Religion-wise Death Count  
7. Place of Death  
8. District-wise Applications  
9. Urban vs Rural Report  
10. Block/Municipality-wise Report  
11. Entered By (Role)  
12. Entered By Institution  
13. Top 10 Immediate Causes of Death  
14. ID Proof Submitted  
15. Occupation of Deceased  
16. Application Status Distribution  
17. Missing First Name Count  
18. Missing ID Proof Count  
19. Manner of Death  

---

 📄 Report Generators

 5. Actual Live Birth Report Generator
Creates a structured report from exported birth application Excel files with:
- SL No.  
- Hospital Name  
- Actual Live Birth Events  
- JMT Entries  
- Pending Entries  
- Completed Entries  

 6. Actual Death Report Generator
Generates summarized death report from death application Excel files with:
- SL No.  
- Hospital Name  
- Actual Death Events  
- JMT Entries  
- Pending Entries  
- Completed Entries  

---

 🔄 File Converter

 7. XLSX → CSV → JSON Converter
A utility to:
- Convert `.xlsx` files to `.csv`
- Then convert `.csv` files to `.json` format for downstream processing or API integration.

---

This toolset provides a robust suite for cleaning, analyzing, and visualizing birth and death application data exported from the Janma-Mrityu Tathya Portal, tailored for streamlined reporting and insight generation.
