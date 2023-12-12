import os
from PyPDF2 import PdfReader, PdfWriter
import pdfrw
from pagelabels import PageLabels, PageLabelScheme

folder_dir = ''
file_name = ''
data_folder_dir = ''

os.chdir(folder_dir)
input_path = f'{file_name}.pdf'
output_path = f'{file_name}_updated.pdf'
data_path = f'{data_folder_dir}/{book_name}.txt'


def roman_to_num(roman):
    ROMAN_TO_NUM = {'i': 1, 'v': 5, 'x': 10, 'l': 50, 'c': 100, 'd': 500, 'm': 1000}
    _roman = roman.strip().lower()
    if not all([n in ROMAN_TO_NUM.keys() for n in _roman]):
        return False
    nums = [ROMAN_TO_NUM[n] for n in _roman]
    num = 0
    for (i, n) in enumerate(nums):
        if i + 1 == len(nums) or nums[i+1] <= n:
            num += n
        else:
            num -= n
    return num


def num_to_roman(num):
    NUM_TO_ROMAN = {1: 'i', 5: 'v', 10: 'x', 50: 'l', 100: 'c', 500: 'd', 1000: 'm'}
    try:
        _num = int(num)
    except ValueError:
        return False
    if _num <= 0 or _num >= 4000:
        return False
    nums = [int(n) for n in str(_num)]
    pow = 1
    roman = ''
    for n in nums[::-1]:
        if n == 0:
            r = ''
        elif n in [1, 2, 3]:
            r = NUM_TO_ROMAN[pow] * n
        elif n == 4:
            r = NUM_TO_ROMAN[pow] + NUM_TO_ROMAN[pow*5]
        elif n == 5:
            r = NUM_TO_ROMAN[pow*5]
        elif n in [6, 7, 8]:
            r = NUM_TO_ROMAN[pow*5] + NUM_TO_ROMAN[pow] * (n - 5)
        elif n == 9:
            r = NUM_TO_ROMAN[pow] + NUM_TO_ROMAN[pow*10]
        roman = r + roman
        pow *= 10
    return roman


def parse_outline(line):
    level = (len(line) - len(line.lstrip())) // 4
    last_blank_index = line.rfind(' ')
    title = line[:last_blank_index].strip()
    page_str = line[last_blank_index+1:].strip()
    # try:
    #     page = int(line[last_blank_index+1:])
    # except ValueError:
    #     raise Exception(f'Error occured in the data file at the following line.\n"{line}"\nThe characters after the last blank are "{line[last_blank_index+1:]}", which is not an integer.')
    return [level, title, page_str]


def parse_label(line):
    labels = [x.split(':') for x in line.split()]
    label_dict = {int(page): label for (page, label) in labels}
    return label_dict


def gen_label_info(label_dict):
    label_info = []
    for (page_no, label) in label_dict.items():
        startpage = page_no - 1
        if label[-1].isnumeric():
            style = 'arabic'
            prefix = label.rstrip('0123456789')
            firstpagenum = int(label.removeprefix(prefix))
        elif label.islower() and roman_to_num(label):
            style = 'roman lowercase'
            prefix = ''
            firstpagenum = roman_to_num(label)
        elif label.isupper() and roman_to_num(label):
            style = 'roman uppercase'
            prefix = ''
            firstpagenum = roman_to_num(label)
        elif label.islower() and len(label) == 1:
            style = 'letters lowercase'
            prefix = ''
            firstpagenum = ord(label) - 96
        elif label.isupper() and len(label) == 1:
            style = 'letters uppercase'
            prefix = ''
            firstpagenum = ord(label) - 64
        else:
            raise Exception(f'The label "({page_no}, {label})" is not in the correct format.')
        label_info.append([startpage, style, prefix, firstpagenum])
    return label_info


def gen_shift(label_dict):
    first_num_label = []
    for page, label in label_dict.items():
        if label.isnumeric() and (not first_num_label or first_num_label[1] > int(label)):
            first_num_label = [page, int(label)]
    return first_num_label[0] - first_num_label[1] - 1


def gen_label_from_page(label_info):
    def page_to_label(n):
        _, first_label = max([(i, label) for i, label in enumerate(label_info) if label[0] + 1 <= n])
        page = n - first_label[0] + first_label[3] - 1
        if first_label[1] == 'arabic':
            return first_label[2] + str(page)
        elif first_label[1] == 'roman lowercase':
            return num_to_roman(page)
        elif first_label[1] == 'roman uppercase':
            return num_to_roman(page).upper()
        elif first_label[1] == 'letters lowercase':
            return chr(page + 96)
        elif first_label[1] == 'letters uppercase':
            return chr(page + 64)
    return page_to_label


def gen_page_from_label(label_info):
    def label_to_page(l):
        x = gen_label_info({1: l})
        l_data = x[0]
        _, first_label = max([(i, label) for i, label in enumerate(label_info) if label[1:2] == l_data[1:2] and label[3] <= l_data[3]])
        return l_data[3] - first_label[3] + first_label[0] + 1
    return label_to_page


def create_data(data_path):
    with open(data_path, 'r') as file:
        lines = file.readlines()
        label_dict = parse_label(lines[0])
        outline_list = [parse_outline(line) for line in lines[1:]]
        label_info = gen_label_info(label_dict)
        shift = gen_shift(label_dict)
        label_to_page = gen_page_from_label(label_info)
    outline = []
    levels = []
    for (i, (level, title, page_str)) in enumerate(outline_list):
        while len(levels) > level:
            levels.pop()
        levels.append(i)
        if len(levels) < 2:
            parent = None
        else:
            parent = levels[-2]
        if page_str[0] == '-':
            try:
                page = int(page_str) + shift
            except ValueError:
                raise Exception(f'Error occured in the data file at the following line.\n"{lines[i+1]}"\nThe page number "{page_str}" is not in the correct format.')
        else:
            try:
                page = label_to_page(page_str) - 1
            except ValueError:
                raise Exception(f'Error occured in the data file at the following line.\n"{lines[i+1]}"\nThe page number "{page_str}" is not in the correct format.')
        outline.append([title, page, parent])
    return (outline, label_info)
            

def add_outline(input_path, output_path, outline):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    item_list = []
    for (title, page, parent_index) in outline:
        if parent_index is None:
            item = writer.add_outline_item(title, page)
            writer.set_page_mode('/UseOutlines')
        else:
            item = writer.add_outline_item(title, page, parent=item_list[parent_index])
            writer.set_page_mode('/UseOutlines')
        item_list.append(item)

    with open(output_path, 'wb') as output_stream:
        writer.write(output_stream)


def add_page_label(input_path, output_path, label_info):
    reader = pdfrw.PdfReader(input_path)
    writer = pdfrw.PdfWriter()
    labels = PageLabels.from_pdf(reader)

    for (startpage, style, prefix, firstpagenum) in label_info:
        scheme = PageLabelScheme(startpage=startpage, style=style, prefix=prefix, firstpagenum=firstpagenum)
        labels.append(scheme)

    labels.write(reader)
    writer.trailer = reader
    with open(output_path, 'wb') as output_stream:
        writer.write(output_stream)


def modify_pdf(input_path, output_path, data_path):
    outline, label_info = create_data(data_path)
    add_outline(input_path, output_path, outline)
    add_page_label(output_path, output_path, label_info)


modify_pdf(input_path, output_path, data_path)
if os.path.exists(output_path):
    os.remove(input_path)
    os.rename(output_path, input_path)
