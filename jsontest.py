import json

# Function to read the configuration from a JSON file
def read_config(file_path):
    try:
        with open(file_path, 'r') as config_file:
            config = json.load(config_file)
        return config
    except FileNotFoundError:
        print("Configuration file not found. Creating a new one.")
        return {}

# Function to write the configuration to a JSON file
def write_config(file_path, config):
    with open(file_path, 'w') as config_file:
        json.dump(config, config_file, indent=4)

# File path for the configuration file
config_file_path = 'config.json'

# Read the configuration from the file
config = read_config(config_file_path)

# Modify the configuration (for example, change the 'active_num' value)
config['active_num'] = 5
config['name'] = "Hello edit"

# Write the updated configuration back to the file
write_config(config_file_path, config)
