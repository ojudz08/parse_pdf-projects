import os, shutil

parent_dir = os.path.abspath(os.path.join(os.path.dirname( __file__ ), '..'))
install_dir = os.path.join(parent_dir, "dist")

# transfer exe to parent directory
if os.path.exists(install_dir):
    for exe_file in os.listdir(install_dir):
        if exe_file[-4:] == ".exe":
            exe_src = os.path.join(install_dir, exe_file)
            exe_dst = os.path.join(parent_dir, exe_file)

            os.replace(exe_src, exe_dst)

buildpath = os.path.join(parent_dir, "build")
specfile = os.path.join(parent_dir, "pdf_merger.spec")

shutil.rmtree(buildpath)
shutil.rmtree(install_dir)
os.remove(specfile)