# PowerPoint Batch Processor

## Overview
The PowerPoint Batch Processor is a Python-based tool designed to automate the modification of PowerPoint (.pptx) files. It processes all presentations in a specified input folder, converting their text to uppercase, and saves the modified files in an output folder.

This tool is particularly useful for batch editing of presentations, saving time and ensuring consistency.

## Features

* Batch Processing: Handles multiple PowerPoint files concurrently for efficient processing.
* Text Modification: Converts all text in the slides to uppercase.
* Customizable Input/Output: Specify folders for input and output files.

## Prerequisites
* Python 3.6+
* Required Python packages:
    * ```python-pptx```

you can install the required package using:
```bash
pip install python-pptx
```

## Project structure
```tree
task2/
├── input   # Folder containing input .pptx files
└── output  # Folder where modified .pptx files are saved
├── main.py # Main script for running the batch processor

```

## Usage
1. Prepare file samples: Place all the .pptx files you want to process in the input folder.
2. Run the Script: Execute the script from the terminal:
```bash
python main.py
```
> **_NOTE:_** you may run `python3 main.py` in case of the above not working
3. Check the Output at `output` folder

