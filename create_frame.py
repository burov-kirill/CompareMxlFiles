import re
import pandas as pd
import unicodedata
def find_row(current_block, current_row, fl):
    strip_text = fl[fl.find(current_block):]
    fl = strip_text
    row_data = re.search('\d+?,\d+?,\d+?,\d+?', strip_text)
    if row_data != None:
        all_numbers = re.findall('\d+', row_data.group(0))
        if all_numbers != None:
            try:
                row = int(all_numbers[0])
            except:
                row = current_row
            else:
                if row == 0:
                    row = current_row
                else:
                    current_row = row
            return row, current_row, fl
    return current_row, current_row, fl

def create_dataframe_from_file(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        fl = file.read()
        str_list = re.findall('\{.*\n\{.*\n?\{?.*?\}?\n?}.*\}[,\d]+', fl)
        d = dict()
        temp_list = []
        current_row = 0
        for item in str_list:
            clear_str = ''.join(item.splitlines())
            count_newline = item.count('\n')
            if count_newline <= 3:
                end_line = item.rfind('}') + 1
                end_item = item[end_line:]
                row, current_row, fl = find_row(item, current_row, fl)
                matched_value = re.findall('\{\"#\",(.*?)\}', clear_str)
                if len(matched_value) > 0:
                    if len(matched_value) == 1:
                        value = matched_value[0]
                    else:
                        value = matched_value[-1]
                    value = unicodedata.normalize("NFKD", value)
                    all_numbers = re.findall('\d+', end_item)
                    if all_numbers != []:
                        if len(all_numbers) == 1:
                            column = int(all_numbers[0])
                        else:
                            if len(all_numbers)>=3:
                                current_row+=1
                                column = int(all_numbers[2])
                            else:
                                continue
                        d.setdefault(row, []).append((column, value))
    result = []
    max_length = len(max(d.items(), key=lambda x: len(x[1]))[1])
    for k, v in d.items():
        correct_values = []
        for col in v:
            correct_values.append(col[1])
        if len(correct_values) < max_length:
            for _ in range(len(correct_values), max_length):
                correct_values.append("")
        result.append(correct_values)

    df = pd.DataFrame(data=result)
    return df
