from docx import Document
from datetime import datetime
import time
import os

def edit_doc(file_path: str, out_path: str, created_datetime:datetime or None=None, last_modified_datetime: datetime or None=None, author: str or None=None):
    document = Document(file_path)
    if created_datetime is not None:
        document.core_properties.created = created_datetime
    if last_modified_datetime is not None:
        document.core_properties.modified = last_modified_datetime
    if author is not None:
        document.core_properties.author = author
    document.save(out_path)
    if last_modified_datetime is not None:
        mod_time = time.mktime(last_modified_datetime.timetuple())
        os.utime(out_path, (mod_time, mod_time))

def process_dir(root_dir: str, out_dir: str, created_datetime:datetime or None=None, last_modified_datetime: datetime or None=None, author: str or None=None):
    if not os.path.isdir(out_dir):
        os.mkdir(out_dir)
    files = os.listdir(root_dir)
    for c_file in files:
        if os.path.isfile(root_dir + c_file):
            process_file(root_dir + c_file, out_dir + c_file)
        else:
            process_dir(root_dir + c_file + '/', out_dir + c_file + '/')

def process_file(root_dir: str, out_dir: str, created_datetime:datetime or None=None, last_modified_datetime: datetime or None=None, author: str or None=None):
    if root_dir.endswith('.doc') or root_dir.endswith('.docx'):
        edit_doc(root_dir, out_dir, created_datetime, last_modified_datetime, author)
c
def generate_datetime_for_docx_custom_logic():
    '''
    If you need custom logic to generate datetime, you can write that logic in this function
    '''
    pass

if __name__ == "__main__":
    in_path = ""
    out_path = ""
    process_dir(in_path, out_path)