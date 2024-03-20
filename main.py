import cv2
from pyzbar.pyzbar import decode
import xlwings as xw

cam = cv2.VideoCapture(0)
cam.set(3, 640)  # Set the width to 640
cam.set(4, 480)  # Set the height to 480
camera = True

while camera:
    success, frame = cam.read()
    if success:
        for barcode in decode(frame):
            decoded_data = barcode.data.decode('utf-8')
            print(barcode.type)
            print(decoded_data)
            wb = xw.Book('/Users/kierangarigan/Desktop/Robotics/ScoutingSheetMaster.xlsm')  # Update the path to your .xlsm file
            sheet = wb.sheets['Sheet1']  # Access the sheet where you want to append data
            next_empty_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
            sheet.range('A' + str(next_empty_row)).value = decoded_data.split(',')
            wb.save()  # Save the workbook
            wb.close()  # Close the workbook
        cv2.imshow('frame', frame)
        if cv2.waitKey(3) & 0xFF == ord('q'):
            break

cam.release()
cv2.destroyAllWindows()
