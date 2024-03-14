import os
import time
import threading
from datetime import datetime

class BackgroundTasks(threading.Thread):
    def __init__(self, path):
        super().__init__()
        self.PATH_TO_SESSIONS_FOLDER = path

    def run(self,*args,**kwargs):
        self.delete_sessions(dead=False)

    def delete_sessions(self, dead):

        while True:
            cur_time = datetime.now()

            for folder in self.PATH_TO_SESSIONS_FOLDER.iterdir():
                if folder.is_dir(): 
                    modification_time = datetime.fromtimestamp(os.path.getmtime(folder))
                    time_difference = cur_time - modification_time
                    if time_difference.total_seconds() > 20:
                        print(f'Removing folder: {folder}')
                        # Remove the folder and its content
                        for root, dirs, files in os.walk(folder, topdown=False):
                            for file in files:
                                os.remove(os.path.join(root, file))
                            for dir in dirs:
                                os.rmdir(os.path.join(root, dir))
                        os.rmdir(folder)
            
            time.sleep(20)