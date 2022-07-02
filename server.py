import os
from flask import Flask, render_template, flash, request, redirect
import pandas as pd

# папка для сохранения загруженных файлов
UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__))

# расширения файлов, которые разрешено загружать
ALLOWED_EXTENSIONS = {"xlsm", "xlsx", "xlsb", "xls"}

app = Flask(__name__)
app.config["SECRET_KEY"] = "super secret key"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


def allowed_file(filename):
    """ Функция проверки допустимого расширения файла """
    return "." in filename and \
           filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def from_excel_to_json(file):
    """
    Функция принимает файл excel и сохраняет в json с таким же названием
    """
    df = pd.read_excel(file, sheet_name=None)
    norm_json = pd.json_normalize(df)
    file_name, file_extension = os.path.splitext(file.filename)
    with open(f"{file_name}.json", "w", encoding='utf-8-sig') as jsonfile:
        norm_json.to_json(jsonfile, orient='records', date_format='iso',
                          force_ascii=False)


@app.route("/upload/excel", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        file_name, file_extension = os.path.splitext(file.filename)
        if os.path.exists(f"{UPLOAD_FOLDER}/{file_name}.json"):
            flash("Файл уже существует")
            return redirect(request.url)

        if file and allowed_file(file.filename):
            flash("Файл принят в обработку")
            from_excel_to_json(file)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
    # app.run()
