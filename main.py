import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class SubFileHandler(FileSystemEventHandler):
    def on_created(self, event):
        try_to_modify_excel('create')

    def on_modified(self, event):
        try_to_modify_excel('modified')

    def on_deleted(self, event):
        try_to_modify_excel('deleted')


def try_to_modify_excel(method):
    print(method)


if __name__ == "__main__":
    observer = Observer()
    event_handler = SubFileHandler()
    path = "G:/Coding/test"
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("关闭脚本")

    finally:
        observer.stop()
        observer.join()
