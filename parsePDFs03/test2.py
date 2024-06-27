import os, sys
import yaml

init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))

config_yaml = os.path.join(init_dir, "config\word_config.yaml")

with open(config_yaml, 'r') as file:
    config = yaml.safe_load(file)

for k, v in config.items():
    print(k, v)