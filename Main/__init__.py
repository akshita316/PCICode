import pathlib
from configparser import ConfigParser as _ConfigParser
import json

_cfg = _ConfigParser()

_path = pathlib.Path("conf/config.ini")
_cfg.read(_path)
location_of_excel = str(_cfg.get('fileMain', 'fileToRead'))
headersToEliminate = json.loads(str((_cfg.get('fileMain', 'headersToEliminate'))))
