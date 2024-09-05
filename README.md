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
- [GUI Usage](#gui-usage)
- [Contributing](#contributing)

## Project Structure

```bash
App/
├── app.py                # ExcelMaster app script
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
   cd App
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

## GUI Usage



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
