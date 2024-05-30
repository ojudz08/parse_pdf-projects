# Delete if the files exist

import os

parent_dir = os.path.abspath(os.path.join(os.path.dirname( __file__ ), '..'))
to_del = ["pdf_merger.exe", "build", "dist"]
#os.listdir(install_dir)

for file in os.listdir(parent_dir):
    if file in to_del:
        del_file = os.path.relpath(os.path.join(parent_dir, file))

        if file[-4:] == ".exe": print(f"Deleting ...{del_file}")
        else: print(f"Deleting ...\{del_file}")

        os.rmdir(os.path.join(parent_dir, file))
        #os.remove()
