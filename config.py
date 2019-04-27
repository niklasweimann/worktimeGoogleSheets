import json
from os import path


class Config:
    def __init__(self, data_path):
        self.data = {"round_to_minutes": 1}
        self.data_path_local = data_path
        self.load_json()

    def load_json(self):
        if not path.isfile(self.data_path_local):
            raise FileNotFoundError("Couldn't find a config.json file")
        with open(self.data_path_local) as json_data_file:
            json_data = json.load(json_data_file)
        self.data = {**self.data, **json_data}

    def get_property(self, name):
        return self.data.get(name)

    def is_valid(self):
        return self.data.get("credentials_path") is not None and self.data.get("spreadsheet_key") is not None \
               and self.data.get("first_data_cell") is not None

    def save_json(self):
        with open(self.data_path_local, "w") as json_data_file:
            json_data_file.write(json.dumps(self.data))
