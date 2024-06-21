# Resume Parser

This program is used to extract information from a PDF resume file. The extracted information can be saved in JSON or Excel format.

## Libraries Used

- `pdfplumber`: It is used to extract text from PDF files.
- `spacy`: Used for text analysis and named entity identification.
- `tkinter`: It is used to create GUI.
- `pandas`: It is used to convert data into DataFrame and create Excel file.
- `xlsxwriter`: Used by `pandas` to write Excel file.

## Details

### Setup

1. Install Python 3.6 or newer.
2. Run the following command in terminal or command prompt to install the required libraries:

```
pip install pdfplumber
pip install spacy
pip install pandas
pip install tk
```

3. Additionally, install the `xlsxwriter` library for Excel output:

```
pip install xlsxwriter
```

4. Download or clone project files.

### Usage

1. In terminal or command prompt, navigate to the project directory.
2. To run the program, use the command:

```
python parser.py
```

3. When the program runs, a GUI screen will open. On this screen, select a PDF resume file and press the "Start" button.
4. Information will be extracted from the selected PDF file and saved in the `output.xlsx` file in Excel format.

## Development

- The project is open source on GitHub. To contribute, please visit the GitHub repository.
- During development, you can modify existing code to add new features or fix bugs.

![program.png](https://github.com/aykutssert/CvParser/blob/main/images/program.png)
![gui_1.png](https://github.com/aykutssert/CvParser/blob/main/images/gui_1.png)
![gui_2.png](https://github.com/aykutssert/CvParser/blob/main/images/gui_2.png)

