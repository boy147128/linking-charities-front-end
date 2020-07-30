# import names
import xlsxwriter
from PIL import ImageGrab
import os
import win32com.client as win32
import random


def create_celeb_excel(start_index, end_index, excel_name):
    # Create an new Excel file and add a worksheet.
    image_celebrity_path = 'C:\\Users\\suprasert.k\\PycharmProjects\\CelebA-HQ-img\\'

    workbook = xlsxwriter.Workbook(excel_name + '.xlsx')
    worksheet = workbook.add_worksheet()

    # Create Header
    header = ['NO.', 'ชื่อบริษัท', 'Card ID', 'ชื่อพนักงาน', 'นามสกุล', 'รหัสพนักงาน', 'ชื่อบัตร',
              'วันที่เริ่มใช้งาน', 'วันหมดอายุ', 'รูป']
    for header_index in range(len(header)):
        worksheet.write(0, header_index, header[header_index])

    worksheet.set_column(1, 1, 14.5)  # Company name width cell
    worksheet.set_column(3, 3, 14.5)  # Name width cell
    worksheet.set_column(4, 4, 14.5)  # Surname width cell
    worksheet.set_column(5, 5, 16)  # Employee No width cell
    worksheet.set_column(7, 7, 14.5)  # Start Date width cell
    worksheet.set_column(8, 8, 14.5)  # End Date width cell
    worksheet.set_column(9, 9, 14.5)  # Image width cell

    company_list = ['EnCo', 'PTT Digital', 'BSA', 'BS']
    employee_no = 9000000 + start_index

    # for row in range(1, 10000):
    row = 1
    for index in range(start_index, end_index):
        worksheet.set_row(row, 105)
        company_name = random.choice(company_list)

        worksheet.write(row, 0, index)
        worksheet.write(row, 1, company_name)
        worksheet.write(row, 2, employee_no)  # Card ID
        worksheet.write(row, 3, 'NAME')  # Name
        worksheet.write(row, 4, 'SURNAME')  # Surname
        worksheet.write(row, 5, company_name + str(employee_no))  # Emp. No
        worksheet.write(row, 6, '-----')  # Card Name
        worksheet.write(row, 7, '01/01/2563')  # Start Date
        worksheet.write(row, 8, '31/12/2563')  # End Date
        # worksheet.insert_image(row, 9, 'python.jpg', {'x_scale': x_scale, 'y_scale': y_scale})
        worksheet.insert_image(row, 9, image_celebrity_path + str(index - 1) + '.jpg', {'x_scale': 0.105, 'y_scale': 0.135})

        employee_no += 1
        row += 1
    # Widen the first column to make the text clearer.
    # worksheet.set_column('A:A', 30)

    workbook.close()


def generate_celeb_excel():
    create_celeb_excel(1, 5001, 'mock_1')
    create_celeb_excel(5001, 10001, 'mock_2')
    create_celeb_excel(10001, 15001, 'mock_3')
    create_celeb_excel(15001, 20001, 'mock_4')
    create_celeb_excel(20001, 25001, 'mock_5')
    create_celeb_excel(25001, 30001, 'mock_6')


def create_digital_excel():
    image_digital_path = 'C:\\Users\\suprasert.k\\PycharmProjects\\Digital-img'

    workbook = xlsxwriter.Workbook('digital-img.xlsx')
    worksheet = workbook.add_worksheet()

    # Create Header
    header = ['NO.', 'ชื่อบริษัท', 'Card ID', 'ชื่อพนักงาน', 'นามสกุล', 'รหัสพนักงาน', 'ชื่อบัตร',
              'วันที่เริ่มใช้งาน', 'วันหมดอายุ', 'รูป']
    for header_index in range(len(header)):
        worksheet.write(0, header_index, header[header_index])

    worksheet.set_column(1, 1, 14.5)  # Company name width cell
    worksheet.set_column(3, 3, 14.5)  # Name width cell
    worksheet.set_column(4, 4, 14.5)  # Surname width cell
    worksheet.set_column(5, 5, 16)  # Employee No width cell
    worksheet.set_column(7, 7, 14.5)  # Start Date width cell
    worksheet.set_column(8, 8, 14.5)  # End Date width cell
    worksheet.set_column(9, 9, 14.5)  # Image width cell

    employee_no = 9500000
    row = 1

    for root, dirs, files in os.walk(image_digital_path):
        for file in files:
            if file.endswith(".jpg"):
                worksheet.set_row(row, 105)
                company_name = 'PTT Digital'
                full_name = root.split('\\')[-1].split(' ')
                image_path = root + '\\' + file

                worksheet.write(row, 0, row)
                worksheet.write(row, 1, company_name)
                worksheet.write(row, 2, employee_no)  # Card ID
                worksheet.write(row, 3, full_name[0])  # Name
                worksheet.write(row, 4, full_name[1] if len(full_name) > 1 else '')  # Surname
                worksheet.write(row, 5, company_name + str(employee_no))  # Emp. No
                worksheet.write(row, 6, '-----')  # Card Name
                worksheet.write(row, 7, '01/01/2563')  # Start Date
                worksheet.write(row, 8, '31/12/2563')  # End Date
                worksheet.insert_image(row, 9, image_path, {'x_scale': 0.10, 'y_scale': 0.13})

                employee_no += 1
                row += 1
            break

    workbook.close()


def find_image():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(r'C:\Users\suprasert.k\PycharmProjects\untitled\images.xlsx')

    for sheet in workbook.Worksheets:
        for i, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith('Picture'):
                shape.Copy()
                image = ImageGrab.grabclipboard()
                # image.save(file_path, 'jpeg')
                image.save('{}.jpg'.format(i + 1), 'jpeg')


# generate_celeb_excel()
create_digital_excel()
# find_image()
