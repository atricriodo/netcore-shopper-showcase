#!/usr/bin/env python3
"""
push_to_google_slides.py
Automates pushing the editable PowerPoint presentation (Netcore_Unified_Commerce_Deck.pptx)
to Google Slides.

Modes supported:
1. Google Drive API (OAuth 2.0 / Service Account): Automatically converts PPTX to native Google Slides.
2. Native Browser Quick-Import: Opens Google Slides / Drive and reveals the file in Finder for instant drag-and-drop.
"""

import os
import sys
import webbrowser
import subprocess

PPTX_PATH = os.path.abspath("Netcore_Unified_Commerce_Deck.pptx")

def try_google_api_upload():
    """Attempts to upload and convert using Google Drive API if credentials exist."""
    cred_files = ["credentials.json", "client_secrets.json", "service_account.json"]
    found_cred = None
    for cf in cred_files:
        if os.path.exists(cf):
            found_cred = cf
            break
            
    if not found_cred:
        return False, "No Google credentials file (credentials.json or service_account.json) found."

    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2 import service_account
        from google_auth_oauthlib.flow import InstalledAppFlow

        SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/presentations']

        if "service_account" in found_cred:
            creds = service_account.Credentials.from_service_account_file(found_cred, scopes=SCOPES)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(found_cred, SCOPES)
            creds = flow.run_local_server(port=0)

        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': 'Netcore - Unified Commerce Experience Platform',
            'mimeType': 'application/vnd.google-apps.presentation'
        }
        media = MediaFileUpload(
            PPTX_PATH,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            resumable=True
        )

        print("Uploading and converting presentation to Google Slides...")
        file = service.files().create(body=file_metadata, media_body=media, fields='id, name, webViewLink').execute()
        slides_url = file.get('webViewLink')
        print(f"\n🎉 Successfully pushed to Google Slides!")
        print(f"🔗 Direct Presentation Link: {slides_url}")
        webbrowser.open(slides_url)
        return True, slides_url
    except Exception as e:
        return False, str(e)

def open_quick_import():
    """Opens Google Drive / Google Slides import and reveals the presentation file."""
    print("=" * 65)
    print("🚀 PUSHING TO GOOGLE SLIDES")
    print("=" * 65)
    print(f"Presentation File: {PPTX_PATH}")
    print(f"File Size: {os.path.getsize(PPTX_PATH):,} bytes\n")

    api_success, msg = try_google_api_upload()
    if api_success:
        return

    print("ℹ️ Google API credentials not configured locally.")
    print("Launching Instant Google Slides / Drive One-Click Upload:")
    print("1. Opening Google Slides in your browser (https://slides.new)...")
    print("2. Opening Google Drive (https://drive.google.com)...")
    print("3. Highlighting the PPTX file in Finder so you can drag & drop it directly!\n")

    # Reveal in macOS Finder
    try:
        subprocess.run(["open", "-R", PPTX_PATH], check=False)
        print("✅ Revealed PPTX file in macOS Finder.")
    except Exception:
        pass

    # Open Google Drive and Google Slides in browser
    try:
        subprocess.run(["open", "https://drive.google.com/drive/my-drive"], check=False)
        subprocess.run(["open", "https://slides.new"], check=False)
        print("✅ Opened Google Drive & Google Slides in browser.")
    except Exception:
        webbrowser.open("https://drive.google.com/drive/my-drive")
        webbrowser.open("https://slides.new")

    print("\n" + "=" * 65)
    print("HOW TO OPEN IN GOOGLE SLIDES (Takes 3 seconds):")
    print("Option A (Fastest): In the Google Drive tab that just opened, drag and drop")
    print("         'Netcore_Unified_Commerce_Deck.pptx' directly into your Drive.")
    print("         Double-click it -> Opens as a native, fully editable Google Slide deck!")
    print("Option B: In the Google Slides tab (slides.new), click:")
    print("         File → Import slides → Upload → Select 'Netcore_Unified_Commerce_Deck.pptx'.")
    print("=" * 65)

if __name__ == "__main__":
    open_quick_import()
