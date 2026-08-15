# -*- coding: utf-8 -*-
"""Export emails from Microsoft Outlook Desktop on Windows."""

import os
import re
import subprocess
import time
import traceback
from datetime import datetime, timedelta

import win32com.client

LOG_FILE = "outlook_errors.txt"


def clean_text(text):
    """Remove common Outlook/Exchange artifacts from plain-text email content."""
    if not text:
        return ""
    text = re.sub(r"</O=[^>]+>", "", text)
    text = re.sub(r"/O=\S+", "", text)
    text = re.sub(r"<[^>]+@[^>]+>", "", text)
    text = re.sub(r"\[cid:[^\]]+\]", "", text)
    text = re.sub(r"_{5,}", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def log_error(context, error):
    """Append an error with context and traceback to the local log file."""
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as file:
            file.write(
                f"\n{'=' * 50}\n"
                f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {context}\n"
                f"Ошибка: {type(error).__name__}: {error}\n"
                f"{traceback.format_exc()}"
            )
    except OSError:
        # Logging should never crash the main export workflow.
        pass


def find_outlook():
    """Find a common Outlook Desktop executable location."""
    paths = [
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\Office16\OUTLOOK.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\Office16\OUTLOOK.EXE"),
    ]

    for path in paths:
        if os.path.exists(path):
            return path
    return None


def ensure_outlook_running():
    """Return the active Outlook COM object, launching Outlook when necessary."""
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        path = find_outlook()
        if path:
            subprocess.Popen([path])
            for _ in range(30):
                time.sleep(1)
                try:
                    return win32com.client.GetActiveObject("Outlook.Application")
                except Exception:
                    continue
    return None


class OutlookSaver:
    """Export messages from the local Outlook Desktop profile."""

    FOLDERS = {
        "inbox": 6,
        "sent": 5,
        "drafts": 16,
        "deleted": 3,
        "junk": 23,
    }

    def __init__(self):
        self.outlook = ensure_outlook_running() or win32com.client.Dispatch(
            "Outlook.Application"
        )
        self.namespace = self.outlook.GetNamespace("MAPI")

        try:
            self.namespace.Logon("", "", False, False)
        except Exception:
            # Outlook may already have an authenticated session.
            pass

        time.sleep(2)
        print(
            "Accounts: "
            + ", ".join(folder.Name for folder in self.namespace.Folders)
        )
        self.show_folders()

    def show_folders(self):
        """Print available top-level mail accounts and their child folders."""
        print("\nFolders:")
        for account in self.namespace.Folders:
            print(f"  [{account.Name}]")
            for folder in account.Folders:
                try:
                    print(f"    - {folder.Name}: {folder.Items.Count}")
                except Exception:
                    continue

    def get_folder(self, name="inbox"):
        """Get one of Outlook's default folders."""
        try:
            return self.namespace.GetDefaultFolder(self.FOLDERS.get(name, 6))
        except Exception:
            return self.namespace.Folders.Item(1).Folders.Item(1)

    def get_emails(self, folder=None, limit=50, days=None, sender=None, subject=None):
        """Collect messages matching optional date, sender and subject filters."""
        folder = folder or self.get_folder()
        items = folder.Items
        count = items.Count
        print(f"Folder {folder.Name}: {count} objects")

        if count == 0:
            return []

        result = []
        cutoff = datetime.now() - timedelta(days=days) if days else None

        for index in range(1, count + 1):
            if limit and len(result) >= limit:
                break

            try:
                message = items.Item(index)
                message_subject = getattr(message, "Subject", None)
                if message_subject is None:
                    continue

                print(f"  -> [{index}] {message_subject[:50]}")

                if cutoff:
                    try:
                        received = getattr(message, "ReceivedTime", None)
                        sent = getattr(message, "SentOn", None)
                        message_time = received or sent
                        if message_time and datetime(
                            message_time.year, message_time.month, message_time.day
                        ) < cutoff:
                            continue
                    except Exception:
                        pass

                if sender:
                    address = str(
                        getattr(message, "SenderEmailAddress", "") or ""
                    )
                    if sender.lower() not in address.lower():
                        continue

                if subject and subject.lower() not in message_subject.lower():
                    continue

                result.append(message)
            except Exception as error:
                log_error(f"Item {index}", error)

        return result

    def save(self, message, path, fmt="txt"):
        """Save one Outlook message and its attachments."""
        safe_subject = re.sub(
            r'[<>:"/\\|?*\r\n]',
            "_",
            str(message.Subject or "NoSubject"),
        )[:60]

        try:
            date = message.ReceivedTime.strftime("%Y%m%d_%H%M")
        except Exception:
            try:
                date = message.SentOn.strftime("%Y%m%d_%H%M")
            except Exception:
                date = datetime.now().strftime("%Y%m%d_%H%M")

        file_path = os.path.join(path, f"{date}_{safe_subject}")

        if fmt == "msg":
            message.SaveAs(file_path + ".msg", 3)
        elif fmt == "txt":
            with open(file_path + ".txt", "w", encoding="utf-8") as file:
                sender = clean_text(getattr(message, "SenderName", "") or "")
                subject = clean_text(str(message.Subject or ""))
                body = clean_text(str(message.Body or ""))
                file.write(f"From: {sender}\nSubject: {subject}\n\n{body}")
        else:
            with open(file_path + ".html", "w", encoding="utf-8") as file:
                body = getattr(message, "HTMLBody", None) or getattr(
                    message, "Body", ""
                )
                file.write(
                    f"<html><meta charset='utf-8'><body>{body}</body></html>"
                )

        try:
            for index in range(1, message.Attachments.Count + 1):
                attachment = message.Attachments.Item(index)
                attachment_dir = file_path + "_att"
                os.makedirs(attachment_dir, exist_ok=True)
                attachment.SaveAsFile(
                    os.path.join(attachment_dir, attachment.FileName)
                )
        except Exception as error:
            log_error(f"Attachments: {message.Subject}", error)

        return file_path

    def bulk_save(self, path=".", folder="inbox", fmt="txt", **filters):
        """Save multiple matching messages to a local directory."""
        os.makedirs(path, exist_ok=True)
        emails = self.get_emails(self.get_folder(folder), **filters)
        print(f"\nFound to save: {len(emails)}")

        saved = 0
        for index, message in enumerate(emails, 1):
            try:
                file_path = self.save(message, path, fmt)
                print(
                    f"  [{index}/{len(emails)}] OK: {os.path.basename(file_path)}"
                )
                saved += 1
            except Exception as error:
                log_error(f"Save: {message.Subject}", error)
                print(f"  [{index}] FAIL: {error}")

        print(f"\nDone! Saved: {saved} to {os.path.abspath(path)}")


def main():
    """Run the interactive command-line interface."""
    print("=" * 40)
    print("   OUTLOOK EMAIL SAVER")
    print("=" * 40)

    try:
        saver = OutlookSaver()
    except Exception as error:
        log_error("Init", error)
        print(f"Outlook connection error: {error}")
        input("Press Enter to exit...")
        return

    while True:
        print("\n1. Inbox  2. Sent  3. Filters  0. Exit")
        choice = input("Choice: ").strip()
        if choice == "0":
            break

        path = input("Output folder [./emails]: ").strip() or "./emails"

        try:
            limit = int(input("Limit [20]: ").strip() or 20)

            if choice == "1":
                saver.bulk_save(path, "inbox", limit=limit)
            elif choice == "2":
                saver.bulk_save(path, "sent", limit=limit)
            elif choice == "3":
                days = input("Days [all]: ").strip()
                sender = input("Sender [all]: ").strip()
                saver.bulk_save(
                    path,
                    "inbox",
                    limit=limit,
                    days=int(days) if days else None,
                    sender=sender or None,
                )
        except Exception as error:
            log_error("Menu", error)
            print(f"Error: {error}")

    input(f"\nError log: {os.path.abspath(LOG_FILE)}\nPress Enter to exit...")


if __name__ == "__main__":
    main()
