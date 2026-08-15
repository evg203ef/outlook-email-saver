# Outlook Email Saver

A small Windows desktop automation tool for exporting emails from a local Microsoft Outlook Desktop profile.

The project was created as a personal learning project while building practical Python and automation skills with AI-assisted development.

## Features

- Connects to Microsoft Outlook Desktop through the Outlook COM/MAPI interface.
- Lists mail accounts and their visible top-level folders.
- Exports messages from Inbox or Sent Items.
- Supports optional filters by sender and date range.
- Saves messages as `.txt`, `.html`, or native Outlook `.msg` files.
- Lets the user choose the export format from the command-line menu.
- Exports message attachments into a separate folder.
- Sanitizes message subjects and attachment names for Windows-safe filenames.
- Avoids overwriting existing exported files by adding a numeric suffix.
- Writes failures and tracebacks to a local `outlook_errors.txt` log.
- Can start Outlook automatically when it is not already running.

## Requirements

- Windows
- Microsoft Outlook Desktop installed and configured with a local profile
- Python 3.9+
- `pywin32`

## Installation

```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
```

## Usage

```bash
python outlook_saver.py
```

The command-line menu provides:

1. Inbox export
2. Sent Items export
3. Filtered Inbox export by age and sender
0. Exit

For each export, the program asks for an output folder, a message limit, and an export format:

- TXT — readable plain-text export
- HTML — preserves the message's HTML body when available
- MSG — native Outlook message format

The default output directory is `./emails`.

## How it works

The application uses `pywin32` to access the Outlook COM object model. It obtains the active Outlook application when possible, starts Outlook when necessary, connects to the MAPI namespace, selects a supported default folder, applies simple filters, and writes each matching message to disk.

Date filtering compares the message timestamp directly with the calculated cutoff. Folder lookup fails explicitly instead of silently falling back to an unrelated folder.

The exporter also checks for attachments and saves them alongside the exported message. Attachment filenames and generated message filenames are sanitized for Windows and existing files are not overwritten.

## Error handling

The project includes lightweight contextual logging. When an individual message or attachment cannot be processed, the error and traceback are written to `outlook_errors.txt` so the rest of the export can continue where possible.

## Security note

This version intentionally does **not** modify Outlook security settings or change Windows Registry values. Earlier development versions experimented with registry changes to suppress Outlook Object Model security prompts; that behavior was removed before publishing the project because changing security settings is inappropriate for a small public portfolio project.

No Outlook password or account credential is stored in the source code. Authentication is provided by the user's existing local Outlook profile.

## AI-assisted development

This project was developed with extensive use of AI coding assistants, including ChatGPT, Cursor, GitHub Copilot and Gemini. AI tools were used to explore implementation approaches, generate and revise code, investigate errors, and iterate on the application. The resulting program was then run and tested locally on a Windows machine.

## Current limitations

- Windows and Outlook Desktop only.
- The implementation relies on the Outlook COM object model rather than the Microsoft Graph API.
- Folder selection currently covers Outlook's main default folders through the CLI; the folder-listing code also shows the accounts and child folders available in the local profile.
- There is currently no automated test suite.
- The UI is intentionally minimal and command-line based.

## Project status

Personal portfolio / learning project. The tool is functional for local Outlook email export and is intended as a practical demonstration of Python, Windows automation, file handling, filtering, error handling, and iterative debugging.
