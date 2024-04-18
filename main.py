import cv2
from pyzbar.pyzbar import decode
import xlwings as xw
import threading
from queue import Queue
import logging
import time

# Setup logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize the queue and lock globally
task_queue = Queue()
excel_lock = threading.Lock()

def worker():
    while True:
        decoded_data = task_queue.get()
        if decoded_data is None:
            # Allows the thread to exit cleanly
            task_queue.task_done()
            break
        try:
            handle_excel(decoded_data)
        except Exception as e:
            logging.error(f"Failed to process data: {e}")
        finally:
            task_queue.task_done()

def handle_excel(decoded_data):
    with excel_lock:
        wb = None
        try:
            wb = xw.Book('**EXCEL SHEET PATH**')
            sheet = wb.sheets['Sheet1']
            next_empty_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
            sheet.range('A' + str(next_empty_row)).value = decoded_data.split(',')
            wb.save()
            logging.info(f"Data written to Excel successfully: {decoded_data}")
        except Exception as e:
            logging.error(f"Error accessing Excel: {e}")
            raise Exception(f"Error accessing Excel: {e}")
        finally:
            if wb:
                try:
                    wb.close()
                except Exception as e:
                    logging.error(f"Error closing Excel: {e}")

# Start worker threads
num_workers = 2  # Adjust based on how fast you need processing and system capability
threads = []
for i in range(num_workers):
    thread = threading.Thread(target=worker)
    thread.start()
    threads.append(thread)

cam = cv2.VideoCapture(0)
cam.set(3, 640)  # Set width
cam.set(4, 480)  # Set height
camera = True

while camera:
    success, frame = cam.read()
    if success:
        for barcode in decode(frame):
            decoded_data = barcode.data.decode('utf-8')
            print(barcode.type)
            print(decoded_data)
            task_queue.put(decoded_data)  
            time.sleep(1)# Add data to the queue
        cv2.imshow('frame', frame)
        if cv2.waitKey(3) & 0xFF == ord('q'):
            break

# Clean up
for _ in threads:
    task_queue.put(None)  # Signal to threads to exit
for thread in threads:
    thread.join()

cam.release()
cv2.destroyAllWindows()
