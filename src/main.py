import os
import io
import sys
from datetime import datetime

import google.auth.transport.requests
from google.oauth2.credentials import Credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/drive.metadata",
]

exportTypesFor = {
    "application/vnd.google-apps.document": [
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
        ("application/rtf", ".rtf"),
        ("application/vnd.oasis.opendocument.text", ".odt"),
        ("text/html", ".html"),
        ("application/pdf", ".pdf"),
        ("text/markdown", ".md"),
        ("application/epub+zip", ".epub"),
        ("text/plain", ".txt"),
    ],
    "application/vnd.google-apps.vid": [("video/mp4", ".mp4")],
    "application/vnd.google-apps.spreadsheet": [
        ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
        ("text/tab-separated-values", ".tsv"),
        ("application/pdf", ".pdf"),
        ("text/csv", ".csv"),
        ("application/vnd.oasis.opendocument.spreadsheet", ".ods"),
    ],
    "application/vnd.google-apps.jam": [("application/pdf", "")],
    "application/vnd.google-apps.script": [("application/vnd.google-apps.script+json", ".json")],
    "application/vnd.google-apps.presentation": [
        ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
        ("application/vnd.oasis.opendocument.presentation", ".odp"),
        ("application/pdf", ".pdf"),
        ("text/plain", ".txt")
    ],
    "application/vnd.google-apps.form": [("application/zip", ".zip")],
    "application/vnd.google-apps.drawing": [
        ("image/png", ".png"),
        ("image/svg+xml", ".svg"),
        ("application/pdf", ".pdf"),
        ("image/jpeg", ".jpg")
    ],
    "application/vnd.google-apps.site": [("text/plain", ".txt")],
    "application/vnd.google-apps.mail-layout": [("text/plain", ".txt")],
    "application/vnd.google-apps.shortcut":  [("text/plain", ".txt")],
}


class GDrivePerms:
    def __init__(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(google.auth.transport.requests.Request())
            else:
                flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            self.service = googleapiclient.discovery.build(
                'drive', 'v3', credentials=creds)
        except HttpError as error:
            print(f'An error occurred: {error}')

    def listDrives(self):
        nextPageTokenD = None
        res = []
        while True:
            results = self.service.drives().list(
                pageToken=nextPageTokenD,
                pageSize=100).execute()
            nextPageToken = results.get("nextPageToken")
            items = results.get('drives', [])
            if not items:
                return res
            res.extend(items)
            if nextPageToken is None:
                break
        res.sort(key=lambda x: x.get("name"))
        return res

    def listRootLevelFiles(self, driveId):
        nextPageToken = None
        res = []
        while True:
            if driveId is None:
                root = self.service.files().get(fileId="root", fields="*").execute()
                self.myDriveOwner = root["owners"][0]["emailAddress"]
                results = self.service.files().list(
                    pageToken=nextPageToken,
                    q="'root' in parents",
                    pageSize=100).execute()
            else:
                results = self.service.files().list(
                    driveId=driveId,
                    includeItemsFromAllDrives=True, corpora="drive", supportsAllDrives=True, spaces="drive",
                    pageToken=nextPageToken,
                    q=f"'{driveId}' in parents",
                    pageSize=100).execute()
            nextPageToken = results.get("nextPageToken")
            items = results.get('files', [])
            if items:
                res.extend(items)
            if nextPageToken is None:
                break
        res.sort(key=lambda x: x.get("name"))
        return res

    def listFilesInDir(self, fileId, fpath):
        nextPageToken = None
        res = []
        while True:
            results = self.service.files().list(
                pageToken=nextPageToken,
                q=f"'{fileId}' in parents",
                pageSize=100,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            ).execute()
            nextPageToken = results.get("nextPageToken")
            items = results.get('files', [])
            if items:
                res.extend(items)
            if nextPageToken is None:
                break
        res.sort(key=lambda x: x.get("name"))
        return res

    def normalize(self, name):
        return name.replace("/", "_")

    def listFiles(self, files, indent, path):
        os.makedirs("./bkup/" + path, exist_ok=True)
        for file in files:
            try:
                isDir = file['mimeType'] == "application/vnd.google-apps.folder"
                fpath = path + self.normalize(file['name'])
                if isDir:
                    fpath += "/"
                if isDir:
                    subFiles = self.listFilesInDir(file["id"], fpath)
                    self.listFiles(subFiles, indent + 3, fpath)
                else:
                    self.handleFile(file, fpath)
            except Exception as e:
                print("Error", e)

    def handleFile(self, file, fpath):
        mt = file["mimeType"]
        if mt.startswith("application/vnd.google-apps."):
            try:
                self.exportG(file, fpath)
            except Exception as e:
                print("ErrorEG", fpath, e)
        else:
            try:
                self.bkupFile(file, fpath)
            except Exception as e:
                print("ErrorRF", fpath, e)

    def bkupFile(self, file, fpath):
        file_id = file["id"]
        file1Info = self.service.files().get(fileId=file_id,
                                             fields="size, modifiedTime",
                                             supportsAllDrives=True,
                                             ).execute()
        file2Path = "./bkup/" + fpath
        if self.probablySame(file1Info, file2Path, False):
            return
        mtime = file1Info["mtime"]
        request = self.service.files().get_media(fileId=file_id)
        fileIO = io.BytesIO()
        downloader = MediaIoBaseDownload(fileIO, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {fpath} {int(status.progress() * 100)}.")
        content = fileIO.getvalue()
        with open(file2Path, "wb") as file2:
            file2.write(content)
        os.utime(file2Path, (mtime, mtime))

    def probablySame(self, file1Info, file2Path, exported):
        file1Mtime = datetime.fromisoformat(file1Info["modifiedTime"])
        file1Secs = file1Mtime.timestamp()
        file1Info["mtime"] = int(file1Secs)

        try:
            file2Info = os.stat(file2Path)
        except:
            # print("missing", file2Path)
            # print("file1", file1Info)
            return False
        if (not exported and int(file1Info["size"]) != file2Info.st_size) or int(file1Secs) != file2Info.st_mtime:
            print("notSame", file2Path)
            print("file1", file1Info)
            print("file2", file2Info)
            return False
        print("isSame", file2Path)
        return True

    def exportG(self, file, fpath):
        file_id = file["id"]
        mimetype = file["mimeType"]
        exportType, fileExt = exportTypesFor[mimetype][0]
        file1Info = self.service.files().get(fileId=file_id,
                                             fields="size, modifiedTime,shortcutDetails",
                                             supportsAllDrives=True,
                                             ).execute()
        if mimetype == "application/vnd.google-apps.shortcut":
            fileSC = self.service.files().get(fileId=file1Info["shortcutDetails"]["targetId"],
                                              fields="id,name,mimeType,size,modifiedTime,shortcutDetails",
                                              supportsAllDrives=True,
                                              ).execute()

            self.handleFile(fileSC, fpath)
            return
        file2Path = "./bkup/" + fpath + fileExt
        if self.probablySame(file1Info, file2Path, True):
            return
        mtime = file1Info["mtime"]
        request = self.service.files().export_media(fileId=file_id, mimeType=exportType)
        fileIO = io.BytesIO()
        downloader = MediaIoBaseDownload(fileIO, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {fpath} {int(status.progress() * 100)}.")
        content = fileIO.getvalue()
        with open(file2Path, "wb") as file2:
            file2.write(content)
        os.utime(file2Path, (mtime, mtime))

    def about(self):
        res = self.service.about().get(fields="exportFormats").execute()
        print(res)


def main():
    namePart = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
    gdp = GDrivePerms()
    # gdp.about()
    print("MyDrive")
    os.makedirs("./bkup", exist_ok=True)
    files = gdp.listRootLevelFiles(None)
    gdp.listFiles(files, 0, "MyDrive/")

    print()
    print()
    print('Geteilte Ablagen:')
    drives = gdp.listDrives()
    for drive in drives:
        drvName = drive["name"]
        if namePart != "" and drvName.lower().find(namePart) == -1:
            continue
        print()
        driveId = drive["id"]
        print(drvName)
        print()
        files = gdp.listRootLevelFiles(driveId)
        os.makedirs("./bkup/" + drvName, exist_ok=True)
        gdp.listFiles(files, 3, drvName + "/")
    print()


if __name__ == '__main__':
    main()
