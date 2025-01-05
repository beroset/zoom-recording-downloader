import base64
import json
import requests
import time
import threading
import datetime
import dateutil.parser as dp
import os
import pathvalidate as path_validate
import platform
import re
import sys
import subprocess
import tkinter as tk
from zoneinfo import ZoneInfo
from tkinter import ttk, filedialog, messagebox
from datetime import date, timedelta


class ZoomConfig:
    """This class reads, writes and holds the configuration data."""

    def __init__(self):
        self.config_file = "~/.zoomvid/zoomvid.conf"
        self.cf = os.path.expanduser(self.config_file)
        self.conf = {
            "OAuth": {
                "account_id": "REPLACE_ME",
                "client_id": "REPLACE_ME",
                "client_secret": "REPLACE_ME",
            },
            "Storage": {"download_dir": "downloads"},
            "Recordings": {
                "timezone": "America/New_York",
                "strftime": "%Y.%m.%d-%H.%M%z",
                "filename": "{meeting_time}-{topic}-{rec_type}-{recording_id}.{file_extension}",
                "folder": "{year}/{month}/{meeting_time}-{topic}",
            },
        }

        if self.exists():
            self.load()
        else:
            self.save()
            self.edit()
            sys.exit(0)

    def edit(self):
        if platform.system() == "Darwin":
            result = subprocess.call(["touch", self.cf])
            result |= subprocess.call(["open", "-a", "TextEdit", self.cf])
        elif platform.system() == "Linux":
            result = subprocess.call(["xdg-open", self.cf])
        return result

    def exists(self):
        return os.path.exists(self.cf)

    def load(self):
        with open(self.cf, encoding="utf-8-sig") as json_file:
            self.conf = json.loads(json_file.read())

    def save(self):
        cfdir = os.path.dirname(self.cf)
        if not os.path.isdir(cfdir):
            os.makedirs(cfdir)

        with open(self.cf, encoding="utf-8-sig", mode="w") as json_file:
            json.dump(self.conf, json_file, indent=4)

    def value(self, section, key):
        return self.conf[section][key]


class ZoomToken:
    """This class gets and holds a Zoom access token"""

    def __init__(self, account_id, client_id, client_secret):
        self.account_id = account_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.expire = time.time()
        self.header = ""
        self.token = ""
        self.last_response = ""

    def format_request(self):
        url = (
            "https://zoom.us/oauth/token",
            "?grant_type=account_credentials",
            f"&account_id={self.account_id}",
        )

        client_cred = f"{self.client_id}:{self.client_secret}"
        client_cred_base64_string = base64.b64encode(
            client_cred.encode("utf-8")
        ).decode("utf-8")

        headers = {
            "Authorization": f"Basic {client_cred_base64_string}",
            "Content-Type": "application/x-www-form-urlencoded",
        }
        return (url, headers)

    def fetch(self, url, headers):
        resp = requests.request("POST", url, headers=headers)
        self.last_response = resp
        response = ""
        if resp.ok:
            response = json.loads(resp.text)
            access_token = response["access_token"]
            self.header = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            self.expire = time.time() + response["expires_in"]
        return response

    def refresh(self):
        if self.expired():
            self.token = self.fetch(*self.format_request())

    def unexpired(self):
        return self.expire - time.time() > 600

    def expired(self):
        return not self.unexpired()


def per_delta(start, end, delta):
    """Generator used to create deltas for recording start and end dates"""
    curr = start
    while curr < end:
        yield curr, min(curr + delta, end)
        curr += delta


def check_file(file_dict):
    """Check to see if the file exists and is the right length.

    Note that because of a strange bug in Zoom, the reported
    length of summary files are 14 bytes shorter than reported.
    Other generated files (e.g. vtt files) can also sometimes
    be reported lengths that are 6 bytes or 2006 bytes short.
    """
    name = file_dict["file_name"]
    reported_len = file_dict["file_size"]
    if not os.path.exists(name):
        return False
    delta = reported_len - os.path.getsize(name)
    if delta == 0:
        return True
    if delta in [6, 14, 2006] and name.endswith((".vtt", ".json")):
        return True
    print(f"MISSING: delta={delta}, reported_len={reported_len}, filename={name}")


class Zoom:
    def __init__(self):
        self.zc = ZoomConfig()
        self.token = ZoomToken(
            self.zc.value("OAuth", "account_id"),
            self.zc.value("OAuth", "client_id"),
            self.zc.value("OAuth", "client_secret"),
        )
        self.last_error = ""
        self.last_response = ""
        self.users = []
        self.meetings = []

    def get_users(self):
        self.token.refresh()
        next_page_token = ""
        more = True
        if self.token.unexpired():
            self.users = []
            while more:
                url = "https://api.zoom.us/v2/users"
                params = {"page_size": 300}
                if next_page_token:
                    params["next_page_token"] = next_page_token
                response = requests.get(url, headers=self.token.header, params=params)
                self.last_response = response
                if response.ok:
                    data = response.json()
                    self.users.extend(data.get("users", []))
                    next_page_token = data.get("next_page_token", "")
                    more = not not next_page_token
                else:
                    self.last_error = response
                    more = False

    def get_meetings_by_user(self, userid, fromdate, todate):
        self.meetings = []
        for start, end in per_delta(
            dp.parse(fromdate), dp.parse(todate), datetime.timedelta(days=30)
        ):
            self.get_meetings_by_user_internal(
                userid, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
            )

    def get_meetings_by_user_internal(self, userid, fromdate, todate):
        self.token.refresh()
        next_page_token = ""
        more = True
        if self.token.unexpired():
            while more:
                url = f"https://api.zoom.us/v2/users/{userid}/recordings"
                params = {"from": f"{fromdate}", "to": f"{todate}", "page_size": 300}
                if next_page_token:
                    params["next_page_token"] = next_page_token
                response = requests.get(url, headers=self.token.header, params=params)
                self.last_response = response
                if response.ok:
                    data = response.json()
                    self.meetings.extend(data.get("meetings", []))
                    next_page_token = data.get("next_page_token", "")
                    more = not not next_page_token
                else:
                    self.last_error = response
                    more = False

    def collect(self, meetings):
        """
        Given a of meetings, extract a list of recordings with:
            recording_type
            file_type
            file_extension
            recording_id
            file_size
            topic (by meeting)
            meeting_time (by meeting)
        """
        return [
            {
                "recording_type": rf["recording_type"],
                "file_type": rf["file_type"],
                "file_extension": rf["file_extension"].lower(),
                "recording_id": rf["id"],
                "file_size": rf["file_size"],
                "download_url": rf["download_url"],
                "topic": item["topic"],
                "start_time": item["start_time"],
            }
            for item in meetings
            for rf in item["recording_files"]
            if rf["file_type"] != ""
        ]

    def synthesize(self, collected_meetings):
        """
        Given the collected meeting from 'collect()', return a list of dictionaries with
            full_filename
            file_size
        """
        return [
            {
                "file_name": self.format_filename(cm),
                "file_size": cm["file_size"],
                "download_url": cm["download_url"],
            }
            for cm in collected_meetings
        ]

    def format_filename(self, params):
        file_extension = params["file_extension"]
        recording_id = params["recording_id"]
        recording_type = params["recording_type"]

        invalid_chars_pattern = r'[<>:"/\\|?*\x00-\x1F]'
        topic = re.sub(invalid_chars_pattern, "", params["topic"])
        rec_type = recording_type.replace("_", " ").title()
        meeting_time_utc = dp.parse(params["start_time"]).replace(
            tzinfo=datetime.timezone.utc
        )
        meeting_time_local = meeting_time_utc.astimezone(
            ZoneInfo(self.zc.value("Recordings", "timezone"))
        )
        year = meeting_time_local.strftime("%Y")
        month = meeting_time_local.strftime("%m")
        day = meeting_time_local.strftime("%d")
        meeting_time = meeting_time_local.strftime(
            self.zc.value("Recordings", "strftime")
        )

        filename = self.zc.value("Recordings", "filename").format(**locals())
        folder = self.zc.value("Recordings", "folder").format(**locals())
        dl_dir = os.sep.join([self.zc.value("Storage", "download_dir"), folder])
        sanitized_download_dir = path_validate.sanitize_filepath(dl_dir)
        sanitized_filename = path_validate.sanitize_filename(filename)
        full_filename = os.sep.join([sanitized_download_dir, sanitized_filename])
        return full_filename

    def download_recording(self, download_url, full_filename, file_size, update):
        dirname = os.path.dirname(full_filename)
        if dirname:
            os.makedirs(dirname, exist_ok=True)

        response = requests.get(download_url, headers=self.token.header, stream=True)

        # total size in bytes.
        total_size = int(response.headers.get("content-length", file_size))
        block_size = 32 * 1024  # 32 Kibibytes

        len_so_far = 0
        try:
            with open(full_filename, "wb") as fd:
                update(len_so_far, total_size)
                for chunk in response.iter_content(block_size):
                    len_so_far += len(chunk)
                    update(len_so_far, total_size)
                    fd.write(chunk)  # write video chunk to disk

            return True

        except Exception as e:
            self.last_error = e
            return False

    def get_file_list(self, startdate, enddate):
        self.get_users()
        self.get_meetings_by_user(self.users[0]["id"], startdate, enddate)
        self.file_list = self.synthesize(self.collect(self.meetings))

    def filter_missing(self):
        return [f for f in self.file_list if not check_file(f)]


class DateDirectorySelector:
    def __init__(self, root, zoom):
        self.root = root
        self.zoom = zoom
        self.root.title("Date and Directory Selector")

        # Start Date Selection
        ttk.Label(root, text="Start Date:").grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        self.start_date = ttk.Entry(root, width=15)
        self.start_date.grid(row=0, column=1, padx=10, pady=10)

        # End Date Selection
        ttk.Label(root, text="End Date:").grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )
        self.end_date = ttk.Entry(root, width=15)
        self.end_date.grid(row=1, column=1, padx=10, pady=10)
        self.end_date.insert(0, date.today().strftime("%Y-%m-%d"))

        # Set initial start date
        initial_end_date = date.today()
        initial_start_date = initial_end_date - timedelta(days=30)
        self.start_date.insert(0, initial_start_date.strftime("%Y-%m-%d"))

        # Directory Selection
        ttk.Label(root, text="Directory:").grid(
            row=2, column=0, padx=10, pady=10, sticky="w"
        )
        self.directory = ttk.Entry(root, width=30)
        self.directory.grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(root, text="Browse", command=self.browse_directory).grid(
            row=2, column=2, padx=10, pady=10
        )
        self.directory.insert(0, zoom.zc.value("Storage", "download_dir"))

        # Buttons
        self.ok_button = ttk.Button(root, text="OK", command=self.ok)
        self.ok_button.grid(row=3, column=1, pady=20, sticky="e")

        self.cancel_button = ttk.Button(root, text="Cancel", command=self.cancel)
        self.cancel_button.grid(row=3, column=2, pady=20, sticky="w")

    def browse_directory(self):
        selected_directory = filedialog.askdirectory()
        if selected_directory:
            self.directory.delete(0, tk.END)
            self.directory.insert(0, selected_directory)

    def ok(self):
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        directory = self.directory.get()

        if not start_date or not end_date or not directory:
            messagebox.showerror("Error", "Please fill in all fields.")
        else:
            self.show_progress_window(start_date, end_date, directory)

    def show_progress_window(self, start_date, end_date, directory):
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Progress")

        ttk.Label(progress_window, text="Processing Files:").grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        self.file_count_label = ttk.Label(progress_window, text="0 of 0")
        self.file_count_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        ttk.Label(progress_window, text="File Name:").grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )
        self.file_name_label = ttk.Label(progress_window, text="None")
        self.file_name_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        ttk.Label(progress_window, text="Current File Size:").grid(
            row=2, column=0, padx=10, pady=10, sticky="w"
        )
        self.file_size_label = ttk.Label(progress_window, text="0 MB")
        self.file_size_label.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        ttk.Label(progress_window, text="File Progress:").grid(
            row=3, column=0, padx=10, pady=10, sticky="w"
        )
        self.file_progress = ttk.Progressbar(
            progress_window, length=300, mode="determinate"
        )
        self.file_progress.grid(row=3, column=1, padx=10, pady=10)

        ttk.Label(progress_window, text="Overall Progress:").grid(
            row=4, column=0, padx=10, pady=10, sticky="w"
        )
        self.overall_progress = ttk.Progressbar(
            progress_window, length=300, mode="determinate"
        )
        self.overall_progress.grid(row=4, column=1, padx=10, pady=10)

        ttk.Label(progress_window, text="Total Size:").grid(
            row=5, column=0, padx=10, pady=10, sticky="w"
        )
        self.total_size_label = ttk.Label(progress_window, text="0 GB")
        self.total_size_label.grid(row=5, column=1, padx=10, pady=10, sticky="w")

        self.progress_ok_button = ttk.Button(
            progress_window,
            text="OK",
            state="disabled",
            command=progress_window.destroy,
        )
        self.progress_ok_button.grid(row=6, column=1, pady=20, sticky="e")

        self.progress_cancel_button = ttk.Button(
            progress_window, text="Cancel", command=progress_window.destroy
        )
        self.progress_cancel_button.grid(row=6, column=2, pady=20, sticky="w")

        # Simulate processing with a thread
        # threading.Thread(target=self.simulate_processing, args=(progress_window,)).start()
        threading.Thread(
            target=self.process,
            args=(
                start_date,
                end_date,
                directory,
                progress_window,
            ),
        ).start()

    def simulate_processing(self, progress_window):
        total_files = 107
        total_size_gb = 1.5  # Example total size in GB
        self.total_size_label.config(text=f"{total_size_gb:.1f} GB")

        for i in range(total_files):
            if not progress_window.winfo_exists():
                break

            # Simulated file name and size
            current_file_name = f"file_{i + 1}.txt"
            current_file_size_mb = (total_size_gb * 1024) / total_files

            # Update file progress
            self.file_count_label.config(text=f"{i + 1} of {total_files}")
            self.file_name_label.config(text=current_file_name)
            self.file_size_label.config(text=f"{current_file_size_mb:.2f} MB")
            self.file_progress["value"] = (i % 100) + 1

            # Update overall progress
            self.overall_progress["value"] = (i / total_files) * 100

            time.sleep(0.05)  # Simulate processing time

        # Task complete
        if progress_window.winfo_exists():
            self.overall_progress["value"] = 100
            self.file_progress["value"] = 100
            self.file_count_label.config(text=f"{total_files} of {total_files}")
            self.file_name_label.config(text="All files processed")
            self.file_size_label.config(text="0 MB")
            self.progress_ok_button["state"] = "normal"

    def update(self, bytes_so_far, total_file_size):
        self.file_progress["value"] = bytes_so_far / total_file_size * 100
        # Update overall progress
        self.overall_progress["value"] = (
            (self.total_bytes_so_far + bytes_so_far) / self.total_size * 100
        )

    def process(self, start_date, end_date, directory, progress_window):
        if self.zoom.zc.value("Storage", "download_dir") != directory:
            self.zoom.zc.conf["Storage"]["download_dir"] = directory
            self.zoom.zc.save()
        self.zoom.get_file_list(start_date, end_date)
        files = self.zoom.filter_missing()
        total_files = len(files)
        self.total_size = sum([f["file_size"] for f in files])
        # print(f"downloading {total_files} files")
        total_size_gb = self.total_size / 2**30
        self.total_bytes_so_far = 0
        self.total_size_label.config(text=f"{total_size_gb:.1f} GB")

        i = 0
        for file in files:
            if not progress_window.winfo_exists():
                break
            current_file_name = file["file_name"]
            current_file_size_mb = file["file_size"] / 2**20

            # Update file progress
            self.file_count_label.config(text=f"{i + 1} of {total_files}")
            self.file_name_label.config(text=os.path.basename(current_file_name)[0:72])
            self.file_size_label.config(text=f"{current_file_size_mb:.2f} MB")
            # print(f"Downloading file {i}")
            self.zoom.download_recording(
                file["download_url"], file["file_name"], file["file_size"], self.update
            )
            self.total_bytes_so_far += file["file_size"]

        # Task complete
        if progress_window.winfo_exists():
            self.overall_progress["value"] = 100
            self.file_progress["value"] = 100
            self.file_count_label.config(text=f"{total_files} of {total_files}")
            self.file_name_label.config(text="All files processed")
            self.file_size_label.config(text="0 MB")
            self.progress_ok_button["state"] = "normal"

    def cancel(self):
        self.root.destroy()


def main():
    zoom = Zoom()
    root = tk.Tk()
    DateDirectorySelector(root, zoom)
    root.mainloop()


if __name__ == "__main__":
    main()
