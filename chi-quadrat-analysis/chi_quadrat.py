import os
import pandas as pd

ROUND_AFTER_POINT = 3
LOWER_LIMIT = 3.8
LIMIT_COLOR = 'cyan'

SHEET_NAME = 'Sheet1'

CURRENT_BOOK = None
TEST_FILE = 'AminoAcids.xlsx'
AMINO_ACIDS = 'ARNDCQEGHILKMFPSTWYV'
RESULT_NAME = 'amino_sequences_result.xlsx'


EXCEL_DATA = []
SEQUENCES_NAMES = []
SEQUENCES_VALUE = []

RESULT = []


def excel_checks(filename):
    if not os.path.exists(filename):
        raise FileNotFoundError(f"File {filename} doesn't exist!")
    if '.xl' not in filename:
        raise ValueError(f"Error! file {filename} isn't a Microsoft Office Excel file!")



def process_excel_data(filename):
    """Reads data from Microsoft Office Excel File to arrays"""

    global DEBUG
    global CURRENT_BOOK
    global SEQUENCES_NAMES
    global RESULT

    excel_checks(filename)
    CURRENT_BOOK = pd.read_excel(filename, sheet_name=SHEET_NAME)
    SEQUENCES_NAMES = CURRENT_BOOK.columns.tolist()[1:]
    for index, row in CURRENT_BOOK.iterrows():
        EXCEL_DATA.append(row.to_list()[1:])
    return analyze_data()

def analyze_data():
    global_result = []
    counter = 0
    SEQUENCES_NAMES[0] = ''
    for i in EXCEL_DATA:
        print(f"Processing for amino acid: {AMINO_ACIDS[counter]}")
        counter+=1

        current_row = i[1:]
        # Add header to current list
        amino_page = []
        amino_page.append(SEQUENCES_NAMES)
        SLICER = 1
        for first_amino in range(len(current_row)-1):
            #add first item of row with calculations results
            amino_cresult = []
            amino_cresult.append(' ')
            row_seqname = SEQUENCES_NAMES[first_amino+1]

            for second_amino in range(len(current_row)):
                if second_amino != first_amino:
                    result = calculate_chi_quadrat(current_row[first_amino], current_row[second_amino])
                    amino_cresult.append(result)
                else:
                    amino_cresult.append(' ')

            amino_cresult = _do_slice(amino_cresult, SLICER)
            amino_cresult[0] = row_seqname
            SLICER+=1

            #add row wih results
            amino_page.append(amino_cresult)

        #Transform current Excel page into DataFrame and color some cells
        cur_result_page = pd.DataFrame(amino_page)
        styler = cur_result_page.style
        styler.applymap(_style_cells)
        global_result.append(styler)

    return global_result

def _do_slice(array, slice):
    for i in range(slice):
        array[i] = ' '
    return array

def _style_cells(val):
    try:
        return 'background-color: %s' % LIMIT_COLOR if float(val) > LOWER_LIMIT else  None
    except:
        return None



def calculate_chi_quadrat(amino_one, amino_two, cor=0):
    """calculates chi quadrat for 2 amino acids """
    delta_amino_one = 100.0 - float(amino_one)
    delta_amino_two = 100.0 - float(amino_two)

    first_difference = abs(amino_one - amino_two)
    second_difference = abs(delta_amino_one - delta_amino_two)
    try:
      answer = first_difference*first_difference/amino_two + second_difference*second_difference/delta_amino_two
      return round(answer, ROUND_AFTER_POINT)
    except:
      global CURRENT_LETTER_DEBUGGER
  
      print(f"\n\nError occured during the operation: Divide by zero!")
      print(f"Current Amino Acid: {CURRENT_LETTER_DEBUGGER}")
      print(f"Amino one: {amino_one}")
      print(f"Amino two: {amino_two}")
      print(f"Delta one: {delta_amino_one}")
      print(f"Delta amino two: {delta_amino_two}\n\n")

      return 0


def save_result(result):
    if len(result) == 0:
        raise ValueError("Result array is empty! There is nothing to save!")
    with pd.ExcelWriter(RESULT_NAME) as writer:
        for current_sheet in range(len(result)):
            result[current_sheet].to_excel(writer, sheet_name=AMINO_ACIDS[current_sheet], header=False, index=False)



def main():

    result = process_excel_data(filename='AminoAcids.xlsx')
    save_result(result)
    print('Done!')

if __name__ == '__main__':
    main()
