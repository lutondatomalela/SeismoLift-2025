# SeismoLift

**SeismoLift** is an open-source tool developed to determine the **seismic categories of elevators** in accordance with **NP EN 1998-1:2009** (*Eurocode 8*) and **EN 81-77* standards.  
Designed specifically for Portugal, it helps engineers, architects, and safety specialists evaluate location-based seismic zones efficiently and reliably.

## 🔍 Features

- Region-specific lookup for **Portugal Continental**, **Madeira**, and **Azores**
- Handles municipalities (concelhos) with duplicate names across regions
- Extracts seismic zone types and acceleration values directly from official data
- Clean CLI interface and modular code
- Generates reports to assist in compliance and documentation

## 🗂 Project Structure
SeismoLift/ 
        
├──── 0_MAIN/SeismoLift.py  (python script)  
├──── 1_IN/Zonas_Sismicas_PT.xlsx  (database)   
├──── 2_OUT/SeismoLift_Report.docx  (gen report)    


## 📥 Official Seismic Data

The file `Zonas_Sismicas_PT.xlsx` contains official seismic classification values for Portuguese municipalities, based on:

📚 **NP EN 1998-1:2009 – Eurocode 8**  
(Anexo Nacional - ANEXO NA.I)

This dataset is publicly referenceable and included for educational and engineering use.


## ⚙️ Usage

1. Clone this repository and open a terminal in the project root.
2. Make sure you have Python 3.7+ and install the dependencies:

- Run the script inside the 0_MAIN directory:
```bash
pip install -r requirements.txt
python SeismoLift.py
```

A .docx report will be saved automatically in the 2_OUT/ folder.

📄 License                          
Distributed under the MIT License. See the LICENSE file for details.

🤝 Contributing                        
Pull requests, bug reports, and feature suggestions are welcome!

📬 Author                   
Created by Engº Lutonda Tomalela —
Bringing seismic clarity to vertical transport.
