from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
from pdf2docx import Converter
from os.path import splitext

#Creates a unique name if a file with the same name exists
def make_unique(source,name):
    filename, extension = splitext(name)
    counter = 1
    while os.path.isfile(source + name) or \
            os.path.isfile(source + name.replace(".pdf", ".docx")):
        name = f"{filename}({str(counter)}){extension}"
        counter += 1
    return name

#Monitors a directory for any .pdf file and converts to .docx
class Convert(FileSystemEventHandler):
    def on_created(self, event):
        source_path = event.src_path #Takes the path of the event from os module
        name = source_path.split("\\")[-1] #extracting the name
        if name.endswith(".pdf"):
            pdf_file = path + name
            if not os.path.isfile(pdf_file.replace(".pdf", ".docx")):
                cv = Converter(pdf_file)
                cv.convert()
                cv.close()
                os.remove(pdf_file)
                print("Converted")
            else:
                new_name = make_unique(path, name)
                os.rename(pdf_file, path + new_name)
                cv = Converter(path + new_name)
                cv.convert()
                cv.close()
                os.remove(path + new_name)
                print("Converted")
                
#Observer start
if __name__ == "__main__":
    event_handler = Convert()
    path = "C:\\Your\\Path\\Here\\"
    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    try:
        print("Monitoring")
        while True:
            pass
    except KeyboardInterrupt:
        observer.stop()
    observer.join()