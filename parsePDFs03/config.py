import os, sys

init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
sys.path.insert(1, os.path.join(init_dir, "scripts"))

import mycredentials, mysystemlib

def config_cred():
    return mysystemlib.get_creds(mycredentials.cred001)