# File-to-Excel Converter and Excel Merger

This repository contains a Python script that provides a graphical user interface (GUI) for converting and merging files to and from Excel format. The program uses the `tkinter`, `pandas`, and `openpyxl` libraries to perform the following functionalities:

## Features

1. **File Converter**: Convert text-based files (e.g., CSV or TXT files) to Excel files.
2. **Excel Merger**: Merge multiple Excel files into a single Excel file.
3. **Excel to Text Converter**: Convert Excel files back to text-based files, generating one file per sheet in the input Excel file.

## Installation

1. Clone this repository to your local machine:
```
git clone https://github.com/andreyrusa/file-to-excel.git
````

2. Install the required dependencies:

```
pip install pandas openpyxl tkinter
```


## Usage

1. Navigate to the directory containing the script:
```
cd file-to-excel
```

2. Run the script:
```
pyinstaller.exe --onfile .\fileToExcel.py
```

3. Run the generated .exe file located in file-to-excel/dist

```
fileToExcel.exe
```

The GUI will appear, and you can use the buttons to select input files for conversion or merging and specify the output directory or file.

## Contributing

Please feel free to submit pull requests for bug fixes, improvements, or new features.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
