import os
from PyPDF2 import PdfWriter, PdfReader, PageObject, Transformation

A4_DIM = (560, 792)

folder_dir = ''
file_name = ''
data_folder_dir = ''

os.chdir(folder_dir)
input_path = f'{file_name}.pdf'
data_path = f'{data_folder_dir}/{file_name}_print.pdf'

def print_pdf_dim(input_path):
    reader = PdfReader(input_path)
    page = reader.pages[0]
    box = page.mediabox
    print(f'left:   {box.left}')
    print(f'right:  {box.right}')
    print(f'bottom: {box.bottom}')
    print(f'top:    {box.top}')

def identify_boxes(page):
    page.cropbox = page.mediabox

def crop_pdf(input_path, output_path, trim, margin=[0, 0, 0, 0], start_page=None, end_page=None):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    left_trim, right_trim, bottom_trim, top_trim = trim
    left_margin, right_margin, bottom_margin, top_margin = margin
    if start_page is None:
        start_page = 0
    else:
        start_page -= 1
    if end_page is None:
        end_page = len(reader.pages)
    pages = reader.pages[start_page:end_page]
    for i, page in enumerate(pages):
        left, right, bottom, top = page.mediabox.left, page.mediabox.right, page.mediabox.bottom, page.mediabox.top
        if i % 2 == 0:
            left += left_trim - left_margin
            right -= right_trim - right_margin
            bottom += bottom_trim - bottom_margin
            top -= top_trim - top_margin
        else:
            left += left_trim - right_margin
            right -= right_trim - left_margin
            bottom += bottom_trim - bottom_margin
            top -= top_trim - top_margin
        page.mediabox.lower_left = (left, bottom)
        page.mediabox.lower_right = (right, bottom)
        page.mediabox.upper_left = (left, top)
        page.mediabox.upper_right = (right, top)
        identify_boxes(page)
        writer.add_page(page)
    with open(output_path, 'wb') as output_stream:
        writer.write(output_stream)

def pdf_to_booklet(input_path, output_path, trim, margin=(9, 9, 0, 0), edge=20, start_page=None, end_page=None, page_list=None, shifts=[0, 0, 0, 0]):
    """
    input_path
        path of the original pdf
    output_path
        path for the output pdf
    trim (left, right, bottom, top)
        amount of trimming from the original pdf
    margin=(9, 9, 0, 0)  (left, right, bottom, top)
        margin on a half-page of the output pdf
    edge=20
        width of the ridge for binding
    start_page=None
        starting page of the output pdf
    end_page=None
        ending page of the output pdf
    page_list=None
        list of page to be put in the output pdf
    shifts=[0, 0, 0, 0] (even_xshift, odd_xshift, even_yshift, odd_yshift)
        shifts for specific printer
    Print instructions:
        Fit to paper | Flip on short edge
    """
    even_xshift, odd_xshift, even_yshift, odd_yshift = shifts
    reader = PdfReader(input_path)
    writer = PdfWriter()
    # Calculate
    w, h = A4_DIM
    ww, hh = reader.pages[0].mediabox.right, reader.pages[0].mediabox.top
    l, r, b, t = trim
    ml, mr, mb, mt = margin
    sx = (h/2-ml-mr-edge) / (ww-l-r)
    sy = (w-mb-mt) / (hh-b-t)
    if sx <= sy:
        s = sx
        txl = -l * s + ml + edge
        tyl = -b * s + (w - (hh-b-t) * s - mb - mt)/2 + mb
        txr = -l * s + ml + h/2
        tyr = -b * s + (w - (hh-b-t) * s - mb - mt)/2 + mb
        print(f'margin l: {ml}, r: {mr}, b: {(w - (hh-b-t) * s - mb - mt)/2 + mb}, t: {(w - (hh-b-t) * s - mb - mt)/2 + mt}')
    else:
        s = sy
        txl = -l * s + (h/2 - (ww-l-r) * s - ml - mr - edge)/2 + ml + edge
        tyl = -b * s + mb
        txr = -l * s + (h/2 - (ww-l-r) * s - ml - mr - edge)/2 + ml + h/2
        tyr = -b * s + mb
        print(f'margin l: {(h/2 - (ww-l-r) * s - ml - mr - edge)/2 + ml}, r: {(h/2 - (ww-l-r) * s - ml - mr - edge)/2 + mr}, b: {mb}, t: {mt}')
    # Rearrange
    if page_list is None:
        if start_page is None:
            start_page = 1
        if end_page is None:
            end_page = len(reader.pages)
        page_list = range(start_page, end_page+1)
    n = len(page_list)
    pages = [reader.pages[i-1] for i in page_list]
    rearranged_pages = []
    for i in range((n+3)//4):
        for k in [2*i, (n+3)//4*4-2*i-1, (n+3)//4*4-2*i-2, 2*i+1]:
            rearranged_pages.append(pages[k]) if k < n else rearranged_pages.append(PageObject.create_blank_page(None, w, h))
    # Merge
    for i in range(len(rearranged_pages) // 2):
        merged_page = PageObject.create_blank_page(None, h, w)
        left_page, right_page = rearranged_pages[i*2], rearranged_pages[i*2+1]
        right_page.mediabox.lower_right = (h, 0)
        identify_boxes(left_page)
        identify_boxes(right_page)
        if i % 2 == 0:
            left_page.add_transformation(Transformation().scale(s).translate(txl+even_xshift, tyl+even_yshift))
            right_page.add_transformation(Transformation().scale(s).translate(txr+even_xshift, tyr+even_yshift))
        else:
            left_page.add_transformation(Transformation().scale(s).translate(txl+odd_xshift, tyl+odd_yshift))
            right_page.add_transformation(Transformation().scale(s).translate(txr+odd_xshift, tyr+odd_yshift))
        merged_page.merge_page(left_page)
        merged_page.merge_page(right_page)
        writer.add_page(merged_page)
    # Write
    with open(output_path, 'wb') as output_stream:
        writer.write(output_stream)

# MATLAB_SHIFT = [-7, 1, -3, 5]
MATLAB_SHIFT = [0, 0, 0, 2]
trim = (50, 50, 50, 50)
crop_pdf(input_path, output_path, trim)
pdf_to_booklet(input_path, output_path, trim, shift=MATLAB_SHIFT)
