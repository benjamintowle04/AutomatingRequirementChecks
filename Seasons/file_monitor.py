import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")
logger = logging.getLogger(__name__)


class Watcher:
    DIRECTORY_TO_WATCH = os.getcwd()  # Using current working directory

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=False)
        self.observer.start()
        logger.info("Watchdog is watching...")
        try:
            while True:
                time.sleep(5)
        except KeyboardInterrupt:
            self.observer.stop()
            logger.info("Observer Stopped")

        self.observer.join()


class Handler(FileSystemEventHandler):
    last_run_time = 0
    run_interval = 10  # Minimum interval in seconds between script runs

    @staticmethod
    def on_any_event(event):
        if event.is_directory:
            return None

        if os.path.basename(event.src_path) == "Shifts.xlsx" or os.path.basename(event.src_path) == "seasons_Guidelines.xlsx":
            Handler.run_script(event)

    @staticmethod
    def run_script(event):
        current_time = time.time()
        if current_time - Handler.last_run_time >= Handler.run_interval:
            logger.info(f"{event.src_path} has been modified or created.")
            subprocess.run(["python", "seasonsGuidelineCheck.py"])
            Handler.last_run_time = current_time  # Update the last run time


if __name__ == "__main__":
    w = Watcher()
    w.run()
