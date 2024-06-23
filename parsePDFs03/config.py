import os, sys

init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
sys.path.insert(1, os.path.join(init_dir, "scripts"))

import mycredentials, mysystemlib

def config01():
    return mysystemlib.get_creds(mycredentials.cred01)

def config02():
    return mysystemlib.get_creds(mycredentials.cred02)