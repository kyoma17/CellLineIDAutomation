import os
import shutil


def PrepTempFolder():
    if not os.path.exists("CellLineTEMP"):
        os.makedirs("CellLineTEMP")

    # Delete all files in the CellLineTEMP folder
    folder = "CellLineTEMP"
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                print("Deleting " + filename)
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                print("Deleting " + filename)
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
