# VBA-challenge

VBA-challenge is a Stock Market Analysis VBA Module for Excel.

[![Excel](https://img.shields.io/badge/Excel_for_MAC-2020-<COLOR>.svg)](https://shields.io/)
![GitHub top language](https://img.shields.io/github/languages/top/OlegRyzhkov2020/VBA-challenge)
[![made-with-Markdown](https://img.shields.io/badge/Made%20with-Markdown-1f425f.svg)](http://commonmark.org)
[![HitCount](http://hits.dwyl.com/OlegRyzhkov2020/VBA-challenge.svg)](http://hits.dwyl.com/OlegRyzhkov2020/VBA-challenge)
![GitHub watchers](https://img.shields.io/github/watchers/OlegRyzhkov2020/VBA-challenge?label=Watch&style=social)
![GitHub followers](https://img.shields.io/github/followers/OlegRyzhkov2020?label=Follow&style=social)

![Excel](images/Excel_Business_Analytics.jpg)

## Installation

The module uses Microsoft Excel. Get it now with a Microsoft 365 subscription

```bash
https://www.microsoft.com/en-us/microsoft-365/excel
```
Please, provide the full path and name for the data file source. By default the provided file below will be used

![Control_Panel](images/Screenshot_ControlPanel.png)

## Usage
As a first step, the module imports data from the Source Workbook and loops through all the stocks for each year over all Worksheets  (see VBA script extract sample below)

```python
'Open Source Workbook
Application.ScreenUpdating = False
Application.DisplayAlerts = False

file_Name = Range("E7")
Set wk = Workbooks.Open(file_Name, ReadOnly:=True)

'Iteration Source Workbook over all sheets
For Each sheet In wk.Worksheets

'Getting Last row number for data on active sheet
    sheet.Activate
    last_Row = Cells(Rows.Count, 1).End(xlUp).Row

```
The summary table outputs the following information:
+ The ticker symbol
+Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
+ The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
+ The total stock volume of the stock.

The solution is able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" per each year.
![Summary](images/Screenshot_Summary.png)

## Contacts
[Find Me on
LinkedIn](https://www.linkedin.com/in/oleg-n-ryzhkov/)
