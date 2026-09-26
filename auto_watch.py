import time
import subprocess
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ReportHandler(FileSystemEventHandler):
    def on_created(self, event):
        self.process(event)

    def on_modified(self, event):
        self.process(event)

    def process(self, event):
        filename = os.path.basename(event.src_path)
        if filename.endswith(".xlsx") and not filename.startswith("~$") and not filename.startswith("Summary_"):
            print(f"\n[+] New sales report detected: {filename}")
            print("[*] Waiting 3 seconds for file write to complete...")
            time.sleep(3)
            
            try:
                print("[*] Running dashboard update & Claude insights...")
                subprocess.run(["python", "generate_dashboard.py"], check=True)
                
                print("[*] Syncing to index.html...")
                subprocess.run(["cmd", "/c", "copy /y Reports\\MMS_Executive_KPI_Dashboard.html index.html"], check=True)
                
                print("[*] Pushing live update to GitHub...")
                subprocess.run(["git", "add", "."], check=True)
                subprocess.run(["git", "commit", "-m", f"Auto sales drop update: {filename}"], check=True)
                subprocess.run(["git", "push", "origin", "main"], check=True)
                print("[✓] SUCCESS: Dashboard updated and published online!\n")
            except Exception as e:
                print(f"[!] Update process error: {e}")

if __name__ == "__main__":
    watch_path = "./reports"
    event_handler = ReportHandler()
    observer = Observer()
    observer.schedule(event_handler, path=watch_path, recursive=False)
    observer.start()
    print("=" * 60)
    print(" [✓] MMS AUTO-WATCHER IS ACTIVE & RUNNING")
    print(" [*] Monitoring folder: ./reports")
    print(" [*] Just drop any new sales excel file into 'reports'...")
    print("=" * 60)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()