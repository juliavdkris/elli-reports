# PRIME automated report generator

## Setup
- `python -m venv .venv`
- `.venv/bin/activate` (bash)  
  `.venv/Scripts/activate` (Windows)
- `pip install -r requirements.txt`


## Usage
- Columns in the spreadsheet that are shown as charts in the document must be named as follows: `num/{ChartName}/{SeriesName}/{ColumnName}`. For example, the column storing the student's Expectancy Beliefs score for week 2 would be named `num/eb/My score/Week2`. Here `eb` is the name of the graph in the Word template, `My score` is the name of the row in the graph's data, and `Week2` is the name of the column in the graph's data.
- The script reads student data from the 'student_identification' sheet of the Excel document. Note that the rows must match the order of the students in the main sheet.
- Make sure that the Word and Excel files used as input match the file path constants `SPREADSHEET` and `TEMPLATE` in `main.py`.
- Now the script can be run with `python main.py`. The generated documents will be saved in the `out` folder.

&nbsp;

Related repository for the library used to manipulate charts in Word documents: https://github.com/juliavdkris/docx-charts
