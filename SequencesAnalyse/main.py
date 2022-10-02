import os
import pandas as pd

# ГИДРОКСИПРОЛИНА НЕТ!
ALLOWED_LETTERS = 'ARNDCQEGHILKMFPSTWYV'


exel_data = {"Letters": list(ALLOWED_LETTERS)}

def normalize_sequence(sequence):
    return ''.join(sequence.splitlines())


def check_sequence(sequence,name):
    sequence = normalize_sequence(sequence)
    current_length = len(sequence)
    result_array = []
    for i in ALLOWED_LETTERS:
        result_array.append((sequence.count(i)/current_length)*100)
    exel_data[name] = result_array


def prepare_exel():
    df = pd.DataFrame(exel_data)
    df.to_excel('./AminoAcids.xlsx')



def read_sequences(filesource):
    with open(filesource) as fs:
        current_line = ''
        current_name = ''
        for line in fs:
            if '>' in line:
                if len(current_line)!=0:
                    check_sequence(current_line, current_name)
                    current_line=''
                first_split =line[line.find('|')+1:]
                current_name = first_split[: first_split.find('|')]
            else:
                current_line+=line


read_sequences('C:\\Users\\Glaster\\Desktop\\SequencesAnalyse\\sequences.txt')
prepare_exel()