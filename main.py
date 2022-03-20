import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


class SubFileHandler(FileSystemEventHandler):
    path_excel = 'F:/Code/test_folder_listener/test.xlsx'
    sheet_name = 'Sheet1'

    def on_created(self, event):
        print(event.event_type)
        work_on_excel(self.path_excel, self.sheet_name, event.event_type)

    '''
    def on_modified(self, event):
        print(event.event_type)
        work_on_excel(self.path_excel, self.sheet_name, event.event_type)
    '''

    def on_deleted(self, event):
        print(event.event_type)
        work_on_excel(self.path_excel, self.sheet_name, event.event_type)


def work_on_excel(path_excel, sheet_name, method):
    try:
        wb = load_workbook(path_excel)
        try:
            ws = wb[sheet_name]

            if method == 'created':
                video_created(ws)

            elif method == 'deleted':
                video_deleted(ws)

            wb.save(path_excel)

        except IOError as msg:
            print(msg)

    except IOError:
        print('无法打开该 Excel 路径: ' + path_excel)


def video_created(work_sheet):
    ws = work_sheet
    column_letter = get_column_letter(ws.max_row)
    ws[column_letter + '1'] = column_letter


def video_modified(work_sheet):
    print(1)


def video_deleted(work_sheet):
    print(1)


if __name__ == "__main__":
    observer = Observer()
    event_handler = SubFileHandler()
    path_test_folder = "F:/Code/test_folder_listener"
    observer.schedule(event_handler, path_test_folder, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("关闭脚本")

    finally:
        observer.stop()
        observer.join()
