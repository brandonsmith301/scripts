import pandas as pd

exchange_abroad_unit_codes = ['units are added here']
PRATO_check = ['units are added here']

class CheckUnits:
    """
    The class CheckUnits checks whether a student is enrolled in at least three units and calculates the total credit points earned.
    """
    
    def __init__(self, file_path: str):
        """
        Initialize the class with the file path.
        
        :param file_path: str, path to the excel file
        """
        self.file = pd.ExcelFile(file_path)
        self.df2 = pd.read_excel(self.file, 'Sheet2')
        self.cols = ['UNIT_1', 'UNIT_2', 'UNIT_3', 'UNIT_4', 'UNIT_5', 'UNIT_6', 'UNIT_7', 'UNIT_8']

    def process_df_exchange(self):
        """
        Process the data frame to get the enrollment status and credit points for the students enrolled in exchange abroad units.
        """
        for col in self.cols:
            self.df2[col] = self.df2[col].apply(str)
            self.df2[col] = self.df2[col].str.slice(stop=7)

        self.unit_dictionary = self.df2.to_dict()

        self.student_id = self.create_id_list(self.unit_dictionary, 'PERSON_ID')
        self.unit_lists = [self.create_unit_list(self.unit_dictionary, col) for col in self.cols]
        self.enrolments = [sum(units) >= 3 for units in zip(*self.unit_lists)]
        self.student_enrolled = list(zip(self.student_id, self.enrolments))
        self.student_enrolled = pd.DataFrame(self.student_enrolled, columns=['Student ID', 'Enrolled?'])

        self.credit_points = [sum(units) * 6 for units in zip(*self.unit_lists)]
        self.student_credit_points = list(zip(self.student_id, self.credit_points))
        self.student_credit_points = pd.DataFrame(self.student_credit_points, columns=['Student ID', 'Credit points'])
        
    def process_df_mgci(self):
        """
        Process the data frame to get the enrollment status and credit points for the students enrolled in MGCI units.
        """
        for col in self.cols:
            self.df2[col] = self.df2[col].apply(str)
            self.df2[col] = self.df2[col].str.slice(stop=7)

        self.unit_dictionary = self.df2.to_dict()

        self.student_id = self.create_id_list(self.unit_dictionary, 'PERSON_ID')
        self.unit_lists = [self.create_unit_list_mgci(self.unit_dictionary, col) for col in self.cols]
        self.enrolments = [sum(units) >= 3 for units in zip(*self.unit_lists)]
        self.student_enrolled = list(zip(self.student_id, self.enrolments))
        self.student_enrolled = pd.DataFrame(self.student_enrolled, columns=['Student ID', 'Enrolled?'])
        self.credit_points = [sum(units) * 6 for units in zip(*self.unit_lists)]
        self.student_credit_points = list(zip(self.student_id, self.credit_points))
        self.student_credit_points = pd.DataFrame(self.student_credit_points, columns=['Student ID', 'Credit points'])

    def create_id_list(self, dictionary, key):
        """
        Create a list of ids.
        
        :param dictionary: dict, the dictionary of the data frame
        :param key: str, the key to extract the id values
        :return: list, the list of ids
        """
        return [d[key] for d in dictionary[key].values()]

    def create_unit_list(self, dictionary, key):
        """
        Create a list of units.
        
        :param dictionary: dict, the dictionary of the data frame
        :param key: str, the key to extract the unit values
        :return: list, the list of units
        """
        return [1 if d[key] in exchange_abroad_unit_codes else 0 for d in dictionary[key].values()]

    def create_unit_list_mgci(self, dictionary, key):
        """
        Create a list of units for MGCI students.
        
        :param dictionary: dict, the dictionary of the data frame
        :param key: str, the key to extract the unit values
        :return: list, the list of units
        """
        return [1 if d[key] in PRATO_check else 0 for d in dictionary[key].values()]

    def create_unit_lists(self):
        """
        Create lists of units and student ids.
        """
        self.process_df_exchange()
        self.process_df_mgci()
        
    def save_to_excel(self, file_path: str):
        """
        Save the results to an excel file.
        
        :param file_path: str, the path to the excel file
        """
        writer = pd.ExcelWriter(file_path)
        self.student_enrolled.to_excel(writer, 'Enrollment', index=False)
        self.student_credit_points.to_excel(writer, 'Credit points', index=False)
        writer.save()

        
# Create an instance of the class and pass the file path to the constructor
checker = CheckUnits("students.xlsx")

# Call the process_df_mgci method to process the data frame for MGCI units
checker.process_df_mgci()

# Access the enrollment status data frame using the student_enrolled property
print(checker.student_enrolled)

# Access the credit points data frame using the student_credit_points property
print(checker.student_credit_points)
