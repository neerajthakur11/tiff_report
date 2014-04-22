'''
    Program to generate the summary of tif files
    Copyright (C) 2014  neerajthakur11

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'''

from os.path import join
from os.path import splitdrive
import os
import time
import datetime
import sys
from PIL import Image
import traceback
import xlsxwriter

SCRIPT_PATH = 'C:/Redist/' 
file_name = 'Job Report.xlsx'
img_file_ext = ['.tif', '.tiff', '.jpg', '.jpeg']
glob_var = 2
use_thumbnail = True

'''
A 1. Date (Date Format)
B 2. DMNo./ Cash
C 3. Thumbnail 
D 4. File name
E 5. Width
F 6. Height
G 7. SQFT (Formula) %5% * %6% (HIDDEN)
H 8. Number of QTY
I 9. Total SQFT (Formula) %7% * %8%
J 10. Rate (Manula fil)
K 11. Amount (fomula) %9% * %10%
L 12. Blank
M 13. PATH (Hidden)
'''

def add_worksheet(workbook):
    global use_thumbnail
    cell_bold_format = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet()
    
    worksheet.freeze_panes(1, 0) #freez the first row

    worksheet.write('A1', 'Modified Date', cell_bold_format)
    worksheet.set_column('A:A', 15)
    
    worksheet.write('B1', 'DMNo. / Cash', cell_bold_format)
    worksheet.set_column('B:B', 10)

    worksheet.write('C1', 'Thumbnail', cell_bold_format)
    if use_thumbnail:
        worksheet.set_column('C:C', 15)
    else:
        worksheet.set_column('C:C', 15, None, {'hidden': 1})


    worksheet.write('D1', 'FileName', cell_bold_format)
    worksheet.set_column('D:D', 35)

    worksheet.write('E1', 'Width', cell_bold_format)
    worksheet.set_column('E:E', 9)

    worksheet.write('F1', 'Height', cell_bold_format)
    worksheet.set_column('F:F', 9)

    worksheet.write('G1', 'SQFT', cell_bold_format)
    worksheet.set_column('G:G', 9, None, {'hidden': 1})

    worksheet.write('H1', 'Qty', cell_bold_format)
    worksheet.set_column('H:H', 9)

    worksheet.write('I1', 'Total SQFT', cell_bold_format)
    worksheet.set_column('I:I', 11)


    worksheet.write('J1', 'Rate', cell_bold_format)
    worksheet.set_column('J:J', 11)
    
    worksheet.write('K1', 'Ammount', cell_bold_format)
    worksheet.set_column('K:K', 11)
    
    # L is left Blank

    worksheet.write('M1', 'Path', cell_bold_format)
    worksheet.set_column('M:M', 90, None, {'hidden': 1})
    #hidden

    return worksheet


def listFiles(report_dir):
    print file_name
    
    try:
        workbook = xlsxwriter.Workbook(join(report_dir, file_name))
        print 'creating file %s'%join(report_dir, file_name)
    except:
        workbook = xlsxwriter.Workbook(join(SCRIPT_PATH, file_name))
    
    print 'report for path %s'%report_dir

    worksheet = add_worksheet(workbook)
    
    for dirname, dirnames, filenames in os.walk(report_dir):
        print 'Walking %s'%str(dirname)
        write_to_worksheet(worksheet, filenames, dirname, report_dir)
    workbook.close()

def write_to_worksheet(worksheet, files, dirname, report_dir):
    global glob_var
    global use_thumbnail
    for each_file in files:
        if not is_filename_valid(each_file):
            continue
        each_file_path = join(dirname, each_file)
        try:
            f_dpi, x_inch, y_inch, div_factor, thumbnail_path, row_height = get_file_details(each_file_path, each_file, dirname)
        except Exception as ex:
            traceback.print_exc()
            f_dpi = x_inch = y_inch = 0
            div_factor = 1
            thumbnail_path = None
        mtime = os.path.getmtime(each_file_path)

        #dt_obj = datetime.datetime.strptime(mtime, "%a %b %d %H:%M:%S %Y") 
        #worksheet.write_datetime('A%d'%glob_var, dt_obj, '%d-%m-%Y')
        worksheet.write('A%d'%glob_var, str(datetime.datetime.fromtimestamp(mtime).strftime("%d-%m-%Y")))
        
        if thumbnail_path != None and use_thumbnail: 
            worksheet.insert_image('C%d'%glob_var, thumbnail_path)
            worksheet.set_row(glob_var-1, max(row_height, 15))
        
        worksheet.write('D%d'%glob_var, each_file)
        
        worksheet.write('E%d'%glob_var, round(x_inch/div_factor, 1))

        worksheet.write('F%d'%glob_var, round(y_inch/div_factor, 1))

        #Formula worksheet.write('G%d'%glob_var, float('%.2f'%f_sqft))
        worksheet.write_formula('G%d'%glob_var,'=E%d*F%d'%(glob_var, glob_var))

        qty_of_job = get_qty_of_job(each_file)
        worksheet.write('H%d'%glob_var, qty_of_job)
        
        
        worksheet.write_formula('I%d'%glob_var,'=G%d*H%d'%(glob_var, glob_var))

        worksheet.write_formula('K%d'%glob_var,'=I%d*J%d'%(glob_var, glob_var))
            
        worksheet.write('M%d'%glob_var, dirname)

        glob_var += 1

        

def get_file_details(file_path, file_name, dirname):
    global use_thumbnail
    THUMB_SIZE = 100, 100
    #Image.DEBUG = True
    fobj = Image.open(file_path)
    xdpi, ydpi = fobj.info['dpi']
    x_px, y_px = fobj.size
    
    xdpi +=1;
    ydpi +=1;

    if x_px % xdpi != 0:
        x_inch = round(x_px / (xdpi * 1.0), 2) 
    else:
        x_inch = x_px / xdpi
    
    if y_px % ydpi != 0:
        y_inch = round(y_px / (ydpi * 1.0), 2)
    else:
        y_inch = y_px / ydpi
    
    
    if 20 < xdpi < 400:
        div_factor = 12
    else:
        div_factor = 1
    
    # generate the thumbnail
    file_name, file_extension = os.path.splitext(file_name)
    
    wlk_drive, wlk_dir = splitdrive(dirname)

    thumbnail_path = join(SCRIPT_PATH, 'thumbnail')
    thumbnail_path = join(thumbnail_path, wlk_dir[1:])
    if not os.path.exists(thumbnail_path):
        os.makedirs(thumbnail_path)
        print 'making dirs %s'%thumbnail_path

    thumbnail_path = join(thumbnail_path, file_name)
    thumbnail_path = thumbnail_path + '.jpg'
    if not os.path.exists(thumbnail_path) and use_thumbnail:
        print 'generating thumbnail for %s'%file_name
        fobj.thumbnail(THUMB_SIZE)
        th_x, th_y = fobj.size
        thumbnail_height = int(th_y * 0.75)
        fobj.save(thumbnail_path, 'JPEG')
    else:
        try:
            thm_obj = Image.open(thumbnail_path)
            th_x, th_y = thm_obj.size
            thumbnail_height = int(th_y * 0.75)
        except:
            thumbnail_height = 10

        print 'skipping generating thumbnail for %s'%file_name
    return xdpi, x_inch, y_inch, div_factor, thumbnail_path, thumbnail_height


def get_qty_of_job(filename):
    qty = 0
    try:
        st_idx = -1
        en_idx = -1
        for i in range(len(filename) - 1, -1, -1):
            if filename[i].isdigit():
                if en_idx == -1:
                    en_idx = i + 1
                    continue
                else:
                    continue
            else:
                if en_idx != -1:
                    st_idx = i + 1
                    break

        qty = int(filename[st_idx:en_idx])
    except:
        traceback.print_exc()
    return qty

def is_filename_valid(filename):
    lfilename = filename.lower()
    file_name, file_extension = os.path.splitext(lfilename)
    if file_extension in img_file_ext:   
        return True
    else:
        return False

if __name__ == '__main__':
    global use_thumbnail
    if len(sys.argv) < 3:
        print 'command usage: processfiles.exe <dir> <thumbnail:YES/NO>'
        exit()
    try:
        if sys.argv[2] == 'YES':
            use_thumbnail = True
        else:
            use_thumbnail = False
    except:
        use_thumbnail = False
    listFiles(sys.argv[1])