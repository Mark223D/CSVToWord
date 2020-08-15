# CSV To Word (Python)

----------------------------------------

This is a python script that will read an csv file, and will input the data of each row into a Word Document using Mergefields.

Inspired by : <https://pbpython.com/python-word-template.html>

## Usage

----------------------------------------

### 1. Install dependencies

    In terminal (MacOSX, Linux) or Powershell (Windows), run `pip install mailmerge` or `pip3 install mailmerge` depending on which version of python you have installed on your computer

### 2. Create MergeFields in Word Document

    - Use <https://pbpython.com/python-word-template.html> as reference to create MergeFields in Word

### 3. Clone repository

### 4. Modify Repo

#### Modify Code Variables

- **template**: Path to **Word** document template
- **csv_path**: Path to **CSV file** containing data to be inputed in template
- **output_path**: Path to **output** folder

#### Modify DataClass

This class contains the fields of each row. Default: Person

##### Default Fields

- **id_number**: Integer (code-generated)
- **surname**: Person's last name
- **name**: Person's first name
- **dob**: Person's date of birth
- **ext_phone**: Person's external phone number
- **local_phone**: Person's local phone number
- **fathers_name**: Person's father's name
- **mothers_name**: Person's mother's name

#### Modify `get_data()` and `fill_template()` functions

- `get_data()`:

  - Based on DataClass
  - Names of columns in CSV File

- `fill_template()`:
  - Based on MergeFields variable names

### 5. Run script

- In terminal (Mac OSX, Linux), or powershell (Windows) navigate to repo folder:
`cd {path}` where `{path}`  is the path to the repo directory

- Once inside the repo directory
 run `python CSVtoWord.py`
 **or**
 run `python3 CSVToWord.py` depending on which version of python you are using.

## Ouput

----------------------------------------

Once you run the script, it will ouput all Word documents inside of the directory defined by `output_path` in the code
