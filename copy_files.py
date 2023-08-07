import os
import filecmp
import shutil

PATH = "/Users/sahilah/Library/CloudStorage/OneDrive-Chalmers/Documents - Chalmers.Northern LEAD Shared Folder/Shared/Lists of Research calls, Special issues, Conferences"
FNAMES = ['Research calls.xlsx', 'Seminars webinars etc.xlsx', 'Special issues - call for papers.xlsx', 'Research conference calls.xlsx']

print(os.listdir(PATH))

for fname in FNAMES:
    f1 = PATH + "/" + fname
    f2 = fname
    if not filecmp.cmp(f1, f2, shallow=True):
        print(f"{fname} – file is modified.")
        shutil.copyfile(f1, f2)
