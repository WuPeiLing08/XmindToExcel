from api.file_conversion import app
from config.setting import SERVER_PORT
import os
import sys
BASE_PATH = os.path.dirname(__file__)
sys.path.insert(0, BASE_PATH)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=SERVER_PORT, debug=True)



