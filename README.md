## Script to automate enrolling students in Ufora groups

Goal: no more carefully clicking 300 checkboxes in the browser. Fill in the spreadsheet and run the script.

## How to use

1- Find your orgUnitId and edit it in script.py

2- Fill in a spreadsheet with the same style as GroupList.xlsx (1st column is student IDs)

3- It will create group categories from column headers, and groups from column values

## Requirements

- Python
https://www.python.org/downloads/

- Chromedriver
https://googlechromelabs.github.io/chrome-for-testing/

- Libraries:
`pip install selenium, pandas, numpy, requests`

## Demo
(outdated)[demo.mp4]
