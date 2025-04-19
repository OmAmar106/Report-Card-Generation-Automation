# Report Card Generation Automation

This project automates the generation of student report cards using **Python** with the following libraries: `pandas`, `openpyxl`, and `win32com`.
It is being used by IIIT Nagpur for generation of Report Cards

## Features

- Extracts student data from an Excel file.
- Reformats the data into a new Excel file.
- For each Student creates a new Excel file format report card.
- Converts the new Excel file into a **PDF report card**.

## Installation 

```bash
>> git clone https://github.com/OmAmar106/Report-Card-Generation-Automation.git
```
```bash
>> cd Report-Card-Generation-Automation
```
```bash
>> cd code
```
```bash
>> pip install -r requirements.txt
```
Navigate to directory containing python scripts
```bash
>> python Scripts/pywin32_postinstall.py -install
```
```bash
>> python main.py
```


