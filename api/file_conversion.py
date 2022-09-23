from flask import Flask
from flask import make_response, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from common.xmind_convert_excel import x_convert_excel

app = Flask(__name__)
ROOT_PATH = os.getcwd()
app.config["ROOT_PATH"] = ROOT_PATH


@app.route("/")
def index():
    return make_response(render_template('index.html'), 200)


@app.route("/convert", methods=["post"])
def convert():
    file_stream = request.files.get("textfield")
    if file_stream is None:
        return "请选择Xmind文件"
    file_name = file_stream.filename
    name, suffix = os.path.splitext(file_name)
    xmind_file = os.path.join(app.config["ROOT_PATH"], file_name)
    file_stream.save(xmind_file)
    convert_file = os.path.join(app.config["ROOT_PATH"], name + '.xlsx')
    r = x_convert_excel(xmind_file, convert_file)
    if r != "ok":
        return r
    try:
        response = make_response(send_file(convert_file, as_attachment=True))
        response.headers["Content-Disposition"] = "attachment; filename={}.xlsx".format(name.encode().decode('latin-1'))
        return response
    except Exception as e:
        return str(e)
    finally:
        if os.path.exists(xmind_file):
            os.remove(xmind_file)
        if os.path.exists(xmind_file):
            os.remove(convert_file)

















