## Table of Contents

- [Table of Contents](#table-of-contents)
- [Introduction](#introduction)
- [Installation](#installation)
- [Usage](#usage)
- [Development](#development)

## Introduction

This script is designed to process and analyze compositions in Microsoft Word (.docx) format. It utilizes various Natural Language Processing (NLP) techniques to extract and categorize cohesive devices, repetitions, substitutions, ellipses, and reiterations within the text. The script also identifies and categorizes synonyms, antonyms, hyponyms, and meronyms for repeated words. Additionally, it finds and counts collocations in the text.
It is designed to analyze text documents, specifically looking for patterns and structures that are common in cohesive texts. It's a powerful tool for educators, researchers, and anyone interested in understanding the linguistic features of texts. The script is written in Python and uses several libraries to perform its tasks. Here's a step-by-step guide on how it works:

## Installation

Before running this script, you need to have Python installed on your computer, you may [download here](https://www.python.org/downloads/).
To install and run this script, follow these steps:

1. **Create a virtual environment (optional but recommended):**
   **Install virtual environment if not available in the user's system:**

   - If you don't have `virtualenv` installed on your system, you can install it using the following command:

   ```
   pip install virtualenv
   ```

   - If you have Python and `virtualenv` installed on your system, you can create a virtual environment for this project using the following command:
     ```
     python3 -m venv env
     ```
   - Activate the virtual environment on your system:
     - Windows:
       ```
       env\Scripts\Activate.bat
       ```
     - macOS/Linux:
       ```
       source env/bin/activate
       ```

2. **Install the required packages:**

   - Install the required packages by running the following command in your terminal:
     ```
     pip install -r requirements.txt
     ```

## Usage

To use this script, follow these steps:

1. **_Update Path:_**

- Inside the script, update the "path/to/directory", you should provide absolute path e.g. "Desktop/Resources/fodler_containing_files"

2. **Run the script:**

   - If you have a virtual environment activated, you can run the script using the following command:
     ```
     python cdv_script.py
     ```

3. **View the results:**
   - The script will print the results for each composition in the specified directory.

## Development

To contribute to this project, follow these steps:

1. **Clone the repository:**

   - Clone the repository using the following command:
     ```
     git clone https://github.com/Chatelo/cohesive_devices_stats.git
     ```

2. Create a virtual environment (optional but recommended):
   Follow the instructions in the Installation section to create a virtual environment for this project.
3. Install the required packages:
   Follow the instructions in the Installation section to install the required packages for this project.
4. Make changes and test the code:
   Make changes to the code and test the changes using the appropriate testing tools and techniques.
5. Submit a pull request:
   Once you have made the desired changes and tested the code, submit a pull request to the main repository.

Contributing
Contributions to this project are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

License
This project is licensed under the MIT License. See the LICENSE file for more information.

**Python Version:**
This script has been tested with Python 3.10 and should work with any Python 3.x version.
