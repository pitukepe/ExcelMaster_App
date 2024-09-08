# ExcelMaster

Automated Excel Analyzer and Enhancer is a Python-based tool designed to streamline data analysis tasks in Excel.
It allows the users to clean data with couple of clicks, create simple pivot tables, plot a variety graphs by choosing pre-loaded column names from their Excel tables, and even integrate and apply Excel formulas without having to understand them deeply.

## Features

- **Data Cleanup**: Perform automated data cleaning operations like removing duplicates or removing rows with a empty value/-s. Output gets saved to a new file in the same directory as the original Excel file, helping to save the original database.
- **Automated Pivot Table Creation**: Automatically generates pivot tables from raw data, saving time and effort in manual operations. Outpus gets saved as a new sheet of the original Excel file.
- **Data Visualization**: Generate charts and graphs to visualize data insights quickly. As of now, graphs are plotted using python pandas library and not directly to Excel.
- **Apply Excel Formulas Programmatically**: Integrate and apply Excel formulas directly to your data files via Python. (// still to be developed)

Providing a GUI that is simple to navigate and easy to use, this tool is ideal for beginner (maybe even all) data analysts looking to automate repetitive tasks in Excel, making data analysis faster and more efficient.

## Table of Contents

- [Project Structure](#project-structure)
- [Technologies Used](#technologies-used)
- [Installation](#installation)
- [Launching the File](#launching-the-file)
- [Contributing](#contributing)
- [GUI Usage example](#gui-usage-example)

## Project Structure

```bash
├── App/
|    └── app.py           # ExcelMaster app script
├── requirments.txt       # Python dependencies
└── README.md
```
- `app.py`: Contains the script, which is the main file that contains the code for the ExcelMaster app.
- `requirements.txt`: Cointains the dependencies required for the script to run.

## Technologies Used

- **python**
- **pandas**
- **numpy**
- **sqlite3**
- **os**
- **tkinter** (plus tkinter's ***ttk***, ***tkinter.filedialog*** and ***tkinter.messagebox***)
- **openpyxl** (plus openpyxl's ***load_workbook*** and openpyxl.utils.dataframe's ***dataframe_to_rows***)
- **matplotlib**

## Installation

To use the Automated Excel Analyzer and Enhancer, follow these steps:

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/pitukepe/ExcelMaster_App.git
   ```
   
2. **Install the Required Dependencies**

   ```bash
   pip install -r requirements.txt
   ```

## Launching the File

### Running the script from the console:
Python Script: Run the script in the script directory to launch the GUI for the ExcelMaster app.
Example:
```bash
cd /Path/To/The/app.py
python3 app.py
```

### Executable file:
You can also create a simple executable file using the pyinstaller package.
```bash
cd /Path/To/The/app.py
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

### MacOs Executable file:
You can also create a simple executable file with .command extension.
```bash
cd /Path/To/The/app.py
touch app.command
nano app.command
```
Then, in the app.command file, write the following:
```bash
cd "$(dirname "$0")"
python3 "$(dirname "$0")/app.py"
echo "App Ran Successfully!"
osascript -e 'tell application "Terminal" to close front window'
```
Save with ^X and close the file. Next, Run:
```bash
chmod +x app.command
```
You can now run the app.command by double tapping on the file.

## Contributing

We welcome contributions to the Automated Excel Analyzer and Enhancer project! If you'd like to contribute, please follow these steps:

1. Fork the repository.
2. Create a new branch.
   ```bash
   git checkout -b YourFeatureName
   ```
3. Commit your changes.
   ```bash
   git commit -a -m 'Add some feature'
   ```
4. Push the branch.
   ```bash
   git push -u origin YourFeatureName
   ```
5. Open a pull request.

## GUI Usage example
### Contents
- [Starting Page](#starting-page)
- [Excel Cleaner Page](#excel-cleaner-page)
- [Excel Pivot table Creator Page](#excel-pivot-table-creator-page)
- [Excel Graph Plotter Page](#excel-graph-plotter-page)
- [Excel Formula Applier Page](#excel-formula-applier-page)

#### Starting Page 
<img width="277" alt="Screenshot 2024-09-05 at 13 50 45" src="https://github.com/user-attachments/assets/920390d0-fd0f-4203-a00b-423a65e03842"></br>
Here, you can choose the tool you want to use.
* `Excel Cleaner` - allows the user to clean an Excel table. (remove duplicates / remove empty values)
* `Excel Pivot Creator` - allows the user to create simple pivot tables, all saved on a new sheet of the original Excel file.
* `Excel Graph Plotter` - allows the user to plot a simple graph based on the specified columns of an Excel table.
* `Excel Formula Applier` - allows the user to input and apply Excel formulas. (❗yet to be developed❗)

#### Excel Cleaner Page
<img width="514" alt="Screenshot 2024-09-05 at 13 51 02" src="https://github.com/user-attachments/assets/3ac3362f-8fb0-46b2-925b-38b06c289fd4"></br>
Here, you can clean an Excel table by following these steps:
1. Choose an Excel file by clicking on the `Choose file` button.
2. If your Excel table has a NUMERICAL index (it gets reset after cleaning), click the check button and then specify it's number like so (e.g. an index in the *first* column would be *1*)</br>
<img width="575" alt="Screenshot 2024-09-05 at 14 25 39" src="https://github.com/user-attachments/assets/3010549e-b190-4cd1-a173-0a9a657b0183"></br>
3. Select a preffered output. `.xlsx` for an Excel file, `.csv` for a CSV file, and `.sqlite` for an SQLite database (you can specify an SQLite table name too)
4. Select a cleaning option. `Drop Duplicates` checkbox drops duplicated rows from the file. `Drop empty rows` checkbox drops rows with empty values (you can further specify `any` for a row with a empty value *anywhere* in the row, or `all` for a row that is completely empty).</br>
<img width="396" alt="Screenshot 2024-09-05 at 14 33 52" src="https://github.com/user-attachments/assets/95501ec4-1d8b-4b3a-baef-aef43c218374"></br>
5. Click the `Clean!` button.
(the cleaned file will be saved at the same directory as the original Excel file, and will have ***_cleaned*** suffix)

#### Excel Pivot table Creator Page
<img width="515" alt="Screenshot 2024-09-05 at 13 51 49" src="https://github.com/user-attachments/assets/692f2077-2b3f-49be-97ea-fde3fa5e185b"></br>
Here, you can create a simple pivot table by following these steps:
1. Choose an Excel file by clicking on the `Choose file` button.
2. Enter the data sheet name (default - *Sheet1*)
3. Specify the *pivot index colum name* (the value to be pivotted) and the *pivot values column name* (the value to be aggregated)
4. Specify the aggregation method for the pivot values column. (options: `sum`(default), `mean`, `max`, `min`, `count`)
5. Click the `Create Pivot Table` button.

Example of an output:
Employee type with the sum of salaries for each employee type.</br>
<img width="169" alt="pivot_example" src="https://github.com/user-attachments/assets/89a0f06c-6b5c-416f-a58c-ef510e1295c6"></br>

#### Excel Graph Plotter Page
<img width="515" alt="Screenshot 2024-09-05 at 13 52 49" src="https://github.com/user-attachments/assets/6b608818-9a14-4fdd-b80b-841d4fba7a56"></br>
Here, you can plot simple graphs by following these steps:
1. Choose an Excel file by clicking on the `Choose file` button.
2. If your Excel table has a NUMERICAL index (it gets reset after cleaning), click the check button and then specify it's number like so (e.g. an index in the *first* column would be *1*)
3. Enter the data sheet name (default - *Sheet1*)
4. Select a Graph Type (options: `box`, `line`, `bar`, `scatter`, `pie`)</br>
With each selection, the page will change according to the necessary inputs for the graph plot.
5. Select values for the graph (for some graphs, `include median` and `include mean` line choices are also available)
6. Click `Plot!` button.

Examples of the choices available for `box` and `bar` graphs:</br>
<img width="515" alt="Screenshot 2024-09-05 at 14 01 26" src="https://github.com/user-attachments/assets/d98ec0af-86ce-4146-a634-bf9c1626b231">
<img width="513" alt="Screenshot 2024-09-05 at 14 01 46" src="https://github.com/user-attachments/assets/c9ed48fe-ed8c-4615-9282-01fd671c8ba6"></br>

Example of a plotted `bar` graph:</br>
<img width="1194" alt="Screenshot 2024-09-05 at 14 08 27" src="https://github.com/user-attachments/assets/c6585c05-69a0-4155-b058-c4a56b05eeec"></br>

#### Excel Formula Applier Page
❗Yet to be developed❗
