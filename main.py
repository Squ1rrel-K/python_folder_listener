import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook


class SubFileHandler(FileSystemEventHandler):
    """
    Those variables will be replaced by user input
    """
    excel_path = 'C:/Users/squ1rrel/Desktop/test/test.xlsx'
    folder_path_tem = 'C:/Users/squ1rrel/Desktop/test/tarkov_tem'
    folder_name_tem = folder_path_tem.split('\\')[-1]
    folder_path = 'C:/Users/squ1rrel/Desktop/test/tarkov'
    folder_name = folder_path.split('\\')[-1]
    sheet_name = 'Sheet1'

    def on_created(self, event):
        work_on_excel(self.excel_path, self.sheet_name, event, self.folder_name_tem, self.folder_name)

    def on_deleted(self, event):
        work_on_excel(self.excel_path, self.sheet_name, event, None, self.folder_name)


def work_on_excel(excel_path, sheet_name, event, folder_tem_name, folder_name):
    try:
        wb = load_workbook(excel_path)
        try:
            ws = wb[sheet_name]

            if event.event_type == 'created':
                work_on_excel_by_video_created(ws, event, folder_tem_name, folder_name)

            elif event.event_type == 'deleted':
                work_on_excel_by_video_deleted(ws, event, folder_name)

            wb.save(excel_path)

        except IOError as msg:
            print(msg)

    except IOError as msg:
        print(msg)


def work_on_excel_by_video_created(work_sheet, event, folder_name_tem, folder_name):
    current_row = work_sheet.max_row + 1
    file_name = event.src_path.split('\\')[-1]
    try:
        work_sheet.cell(column=1, row=current_row, value="{0}".format(file_name)).hyperlink = event.src_path
        work_sheet.cell(column=4, row=current_row, value="No")
        print('Successful write: ' + '[' + file_name + ']' + ' From: ' + folder_name_tem + ' To ' + folder_name)

    except IOError as msg:
        print(msg)


def work_on_excel_by_video_deleted(work_sheet, event, folder_name):
    try:
        file_name = event.src_path.split('\\')[-1]

        for row in range(2, work_sheet.max_row + 1):
            if work_sheet.cell(column=1, row=row).internal_value is not None:
                internal_value = work_sheet.cell(column=1, row=row).internal_value

                if internal_value == file_name:
                    work_sheet.delete_rows(idx=row)
                    for row_sub in range(row, work_sheet.max_row + 1):
                        work_sheet.cell(column=1, row=row_sub).hyperlink = \
                            work_sheet.cell(column=1, row=row_sub).hyperlink.target

        print('Successful delete: ' + '[' + file_name + ']' + ' From: ' + folder_name)

    except IOError as msg:
        print(msg)


if __name__ == "__main__":
    print('Welcome to Python_folder_listener!')
    observer = Observer()
    event_handler = SubFileHandler()
    observer.schedule(event_handler, event_handler.folder_path, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("关闭脚本")

    finally:
        observer.stop()
        observer.join()
