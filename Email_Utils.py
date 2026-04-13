# -*- coding: utf-8 -*-

"""
email_operations.py

Cross-platform Outlook email sender for the Trade Reconciliation project.

Supports:
- Windows Outlook via COM (pywin32)
- macOS Outlook via AppleScript (osascript)

Attaches Tableau-ready output CSVs and includes:
- execution time
- executed user
- total trades
- match rate
"""

from __future__ import annotations

import getpass
import os
import platform
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd


class EmailOperations:
    def __init__(
        self,
        to_email: str = "kwavhal@stevens.edu",
        project_name: str = "Trade Reconciliation & Exception Management Platform",
        sender_name: str = "Komal",
        base_dir: str | None = None,
    ) -> None:
        self.to_email = to_email
        self.project_name = project_name
        self.sender_name = sender_name
        self.base_dir = Path(base_dir) if base_dir else Path(__file__).resolve().parent
        self.output_folder = self.base_dir / "output"

        self.attachments = [
            "reconciliation_detail.csv",
            "exceptions_detail.csv",
            "kpi_summary.csv",
            "status_summary.csv",
            "severity_summary.csv",
            "symbol_breaks.csv",
        ]

    # ============================================================
    # Helpers
    # ============================================================

    def log(self, message: str) -> None:
        print(message)

    def get_execution_user(self) -> str:
        try:
            return getpass.getuser()
        except Exception:
            return "Unknown User"

    def get_execution_time(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def get_existing_attachments(self) -> list[Path]:
        files: list[Path] = []
        for file_name in self.attachments:
            file_path = self.output_folder / file_name
            if file_path.exists():
                files.append(file_path)
            else:
                self.log(f"[WARN] Attachment not found, skipping: {file_path}")
        return files

    def build_metrics(self) -> tuple[int, int, float]:
        recon_file = self.output_folder / "reconciliation_detail.csv"
        if not recon_file.exists():
            raise FileNotFoundError(
                f"reconciliation_detail.csv not found: {recon_file}"
            )

        df = pd.read_csv(recon_file)

        total = len(df)
        matched = int((df["recon_status"] == "MATCHED").sum())
        match_rate = round((matched / total) * 100, 2) if total > 0 else 0.0

        return total, matched, match_rate

    def build_email_subject(self) -> str:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return f"{self.project_name} | Demo Output ({timestamp})"

    def build_email_body(self) -> str:
        total, matched, match_rate = self.build_metrics()

        execution_time = self.get_execution_time()
        user = self.get_execution_user()
        
        
        return f"""
        Hi Shiv and Bharath,
        
        I wanted to share the outputs from my Trade Reconciliation & Exception Management Platform project.
        
        Execution Details:
        - Executed By: {user}
        - Execution Time: {execution_time}
        - Total Trades Processed: {total}
        - Matched Trades: {matched}
        - Match Rate: {match_rate}%
        
        Project Highlights:
        - Automated reconciliation of internal trade records against external market data
        - Exception classification and severity tagging
        - Tableau-ready reporting outputs
        
        Attached are the output files for review.
        
        Thanks,
        Komal
        """
        
        
        # return f"""
        # <p>Hi,</p>

        # <p>I wanted to share my <b>{self.project_name}</b> project outputs.</p>

        # <p><b>Execution Details:</b></p>
        # <ul>
        #     <li>Executed By: {user}</li>
        #     <li>Execution Time: {execution_time}</li>
        #     <li>Total Trades Processed: {total}</li>
        #     <li>Matched Trades: {matched}</li>
        #     <li>Match Rate: {match_rate}%</li>
        # </ul>

        # <p><b>Highlights:</b></p>
        # <ul>
        #     <li>Automated trade reconciliation using market data inputs</li>
        #     <li>Exception classification and severity tagging</li>
        #     <li>Generated Tableau-ready reporting outputs</li>
        # </ul>

        # <p>Attached are the output files for review.</p>

        # <p>Thanks,<br>{self.sender_name}</p>
        # """

    def _escape_applescript(self, text: str) -> str:
        return (
            text.replace("\\", "\\\\")
            .replace('"', '\\"')
            .replace("\n", "\\n")
        )

    # ============================================================
    # Windows Outlook
    # ============================================================

    def send_windows(self) -> None:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError(
                "pywin32 is required on Windows for Outlook COM automation. "
                f"Install with 'pip install pywin32'. Original error: {exc}"
            ) from exc

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = self.to_email
        mail.Subject = self.build_email_subject()
        mail.HTMLBody = self.build_email_body()

        for file_path in self.get_existing_attachments():
            mail.Attachments.Add(str(file_path))

        mail.Send()
        self.log("✅ Email sent via Windows Outlook")

    # ============================================================
    # macOS Outlook
    # ============================================================

    def send_mac(self) -> None:
   
        if not shutil.which("osascript"):
            raise RuntimeError("osascript not found on this macOS system.")
    
        subject = self._escape_applescript(self.build_email_subject())
        body = self._escape_applescript(self.build_email_body())
        to_email = self._escape_applescript(self.to_email)
    
    
    
        attachment_lines = []
        for file_path in self.get_existing_attachments():
            attachment_lines.append(
                f'set aFile to POSIX file "{self._escape_applescript(str(file_path))}"\n'
                f'make new attachment with properties {{file:aFile}} at newMessage'
            )
    
        attachment_block = "\n".join(attachment_lines)
    
        apple_script = f'''
        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
            make new recipient at end of to recipients of newMessage with properties {{email address:{{address:"{to_email}"}}}}
            {attachment_block}
            send newMessage
        end tell
        '''
    
        try:
            subprocess.run(
                ["osascript", "-e", apple_script],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            stderr = (exc.stderr or "").strip()
            raise RuntimeError(
                "Failed to send via Outlook on macOS.\n"
                f"AppleScript error:\n{stderr}"
            ) from exc
    
        self.log("✅ Email sent via Mac Outlook")

    # ============================================================
    # Main dispatcher
    # ============================================================

    def send_email(self) -> None:
        os_name = platform.system().lower()

        if os_name == "windows":
            self.send_windows()
            return

        if os_name == "darwin":
            self.send_mac()
            return

        raise RuntimeError(
            f"Unsupported OS for Outlook-direct sending: {platform.system()}"
        )


# if __name__ == "__main__":
#     emailer = EmailOperations(
#         to_email="kwavhal@stevens.edu",
#         sender_name="Komal",
#     )
#     emailer.send_email()


















































# import platform
# import subprocess
# import shutil
# from datetime import datetime
# import os 


# class ExcelOperations:
       
        
#     # ===========================
#     # CONFIG
#     # ===========================
    
#     TO_EMAIL = "kwavhal@stevens.edu"
    
#     BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
#     OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
    
#     ATTACHMENTS = [
#         "reconciliation_detail.csv",
#         "exceptions_detail.csv",
#         "kpi_summary.csv",
#         "status_summary.csv",
#         "severity_summary.csv",
#         "symbol_breaks.csv"
#     ]
    
#     # ===========================
#     # BUILD DYNAMIC METRICS
#     # ===========================
    
#     def build_metrics():
#         df = pd.read_csv(os.path.join(OUTPUT_FOLDER, "reconciliation_detail.csv"))
    
#         total = len(df)
#         matched = (df["recon_status"] == "MATCHED").sum()
#         match_rate = round((matched / total) * 100, 2)
    
#         return total, match_rate
    
#     # ===========================
#     # EMAIL BODY
#     # ===========================
    
#     def build_email_body():
#         total, match_rate = build_metrics()
    
#         execution_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#         user = os.getlogin()
    
#         return f"""
#         <p>Hi,</p>
    
#         <p>I wanted to share a Trade Reconciliation & Exception Management project I recently implemented.</p>
    
#         <p><b>Execution Details:</b></p>
#         <ul>
#             <li>Executed By: {user}</li>
#             <li>Execution Time: {execution_time}</li>
#             <li>Total Trades: {total}</li>
#             <li>Match Rate: {match_rate}%</li>
#         </ul>
    
#         <p><b>Highlights:</b></p>
#         <ul>
#             <li>Automated trade reconciliation using market APIs</li>
#             <li>Exception classification & severity tagging</li>
#             <li>Tableau-ready reporting outputs</li>
#         </ul>
    
#         <p>Attached are the output files for review.</p>
    
#         <p>Thanks,<br>Komal</p>
#         """

#     # ===========================
#     # WINDOWS OUTLOOK
#     # ===========================
    
#     def send_windows(self, OUTPUT_FOLDER):
#         import win32com.client
    
#         outlook = win32com.client.Dispatch("Outlook.Application")
#         mail = outlook.CreateItem(0)
    
#         mail.To = TO_EMAIL
#         mail.Subject = "Trade Reconciliation Project Demo"
#         mail.HTMLBody = build_email_body()
    
#         # Attach files
#         for file in ATTACHMENTS:
#             file_path = os.path.join(OUTPUT_FOLDER, file)
#             if os.path.exists(file_path):
#                 mail.Attachments.Add(file_path)
    
#         mail.Send()
#         print("✅ Email sent via Windows Outlook")
    
#     # ===========================
#     # MAC OUTLOOK (APPLESCRIPT)
#     # ===========================
    
#     def send_mac(self, OUTPUT_FOLDER):
#         if not shutil.which("osascript"):
#             raise RuntimeError("osascript not found")
    
#         body = build_email_body()
    
#         attachment_commands = ""
#         for file in ATTACHMENTS:
#             file_path = os.path.join(OUTPUT_FOLDER, file)
#             if os.path.exists(file_path):
#                 attachment_commands += f'''
#                 make new attachment with properties {{file name: POSIX file "{file_path}"}} at after the last paragraph
#                 '''
    
#         script = f'''
#         tell application "Microsoft Outlook"
#             set newMessage to make new outgoing message with properties {{subject:"Trade Reconciliation Project Demo", content:"{body}"}}
#             make new recipient at end of to recipients of newMessage with properties {{email address:{{address:"{TO_EMAIL}"}}}}
#             tell newMessage
#                 {attachment_commands}
#             end tell
#             send newMessage
#         end tell
#         '''
    
#         subprocess.run(["osascript", "-e", script])
#         print("✅ Email sent via Mac Outlook")
