import os
from PyPDF2 import PdfReader

folder_name = ''
file_name = ''

os.chdir(folder_name)
input_path = f'Other Topics/{file_name}.pdf'
data_path = f'Modify PDF/{file_name}.txt'

shift = 1 # at which page is page 1

def print_outline(reader, outline, shift, ljust_amt=150, indent=0):
    lines = []
    for outline_item in outline:
        if isinstance(outline_item, list):
            _lines = print_outline(reader, outline_item, shift, indent=indent+1)
            lines += _lines
        else:
            output_title = ' ' * 4 * indent + outline_item.title.replace('\x00', '')
            output_page_number = reader.get_destination_page_number(outline_item) - shift + 2
            line = output_title.ljust(ljust_amt) + str(output_page_number) + '\n'
            lines.append(line)
    return lines

reader = PdfReader(input_path)
lines = print_outline(reader, reader.outline, shift)
with open(data_path, 'w') as data_stream:
    data_stream.writelines(lines)

