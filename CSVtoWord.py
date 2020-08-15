from __future__ import print_function
from mailmerge import MailMerge
import csv


template = "template.docx" # Path to word document containing MergeFields
csv_path = "data.csv"      # Path to csv file containing data
output_path = "output"     # Path to output folder that will contain generated word documents

data = []

# EDIT STARTS AT THE FOLLOWING LINE
# Modify the following class based on the data inside the CSV file.
# Default: Person Class
class DataClass:
    def __init__(self, id_number, surname, name, dob, ext_phone, father_name, mother_name, local_phone):
        self.id_number = id_number
        self.surname = surname
        self.name = name
        self.dob = dob
        self.ext_phone = ext_phone
        self.local_phone = local_phone
        self.father_name = father_name
        self.mother_name = mother_name
# EDIT ENDS AT THE PREVIOUS LINE

# The following function will get the data from the csv file and input it into the data array
def get_data():
    # Open csfv file
    with open(csv_path) as f:
        # Convert CSV file to Array of Python dictionaries
        data_arr_dict = [{k: str(v) for k, v in row.items()} for row in csv.DictReader(f, skipinitialspace=True)]
        # Convert Dictionaries to DataClass Objectsand input into data Array  
        [data.append(
                    # EDIT STARTS AT FOLLOWING LINE
                    # Edit the following lines based on the fields in the DataClass(i.e. id_number, surname, ....) and the CSV column names (i.e. 'Surname', 'Other Name'...)
                    DataClass(id_number=str(len(people)),
                                surname=d['Surname'],
                                name=d['Other Name'],
                                dob=d['Date of Birth'],
                                ext_phone=d['External Phone Number'],
                                father_name=d['Father\'s Name'],
                                mother_name=d['Mother\'s Name'],
                                local_phone=d['Local Phone Number']
                    # EDIT ENDS AT PREVIOUS LINE
                            )
                )
        for d in data_arr_dict]

# The following function will generate word documents and fill each with the data from the data array based on the provided mergefields
def fill_template():
    # Loop through all DataClass object inside of data array
    for d in data:
        # Create new word document from template 
        document = MailMerge(template)
        # EDIT STARTS AT FOLLWING LINE

        # The following line contains paramaters named after mergefields created in word document
        document.merge(
                    Surname=d.surname.upper(),
                    ExtPhone=d.ghana_phone.upper(),
                    DOB=d.dob,
                    FatherName=d.father_name.upper(),
                    MotherName=d.mother_name.upper(),
                    IDNUMBER=d.id_number,
                    LocalPhone=d.local_phone,
                    Forename=d.name.upper())
        # Optional: The following line sets the name of each word document created
        file_name =  d.id_number + '-' +d.surname.upper()+'-'+ d.name.replace('/', '').upper()
        
        # EDIT ENDS AT PREVIOUS LINE
        document.write(output_path+'/'+file_name+'.docx')


# The follwoing function will run when the program will be run
def run():
    # Calling get_data() function to get data from csv in data Array
    get_data()
    # Calling fill_template() function and passsing data in order to produce & output word documents into output folder from template.
    fill_template()

# Calling run() function
run()