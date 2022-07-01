import os
from flask import Flask, render_template, flash, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename

# папка для сохранения загруженных файлов
UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__))
# расширения файлов, которые разрешено загружать
ALLOWED_EXTENSIONS = {'xlsm', 'xlsx'}

app = Flask(__name__)
# app.secret_key = 'super secret key'
app.config['SECRET_KEY'] = 'super secret key'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    """ Функция проверки расширения файла """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/uploads/<name>')
def download_file(name):
    return send_from_directory(app.config["UPLOAD_FOLDER"], name)

@app.route("/upload/excel", methods=['GET', 'POST'])
def upload_file():
    # return "файл принят в обработку"
    if request.method == 'POST':
        # проверим, передается ли в запросе файл 
        if 'file' not in request.files:
            # После перенаправления на страницу загрузки
            # покажем сообщение пользователю 
            flash('Не могу прочитать файл')
            return redirect(request.url)
        file = request.files['file']
        # Если файл не выбран, то браузер может
        # отправить пустой файл без имени.
        if file.filename == '':
            flash('Нет выбранного файла')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # безопасно извлекаем оригинальное имя файла
            filename = secure_filename(file.filename)
            # сохраняем файл
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # если все прошло успешно, то перенаправляем  
            # на функцию-представление `download_file` 
            # для скачивания файла
            # return redirect(url_for('download_file', name=filename))

    return render_template('index.html')
    # return '''
    # <!doctype html>
    # <title>Загрузить новый файл</title>
    # {% with messages = get_flashed_messages() %}
    #   {% if messages %}
    #     <ul class=flashes>
    #     {% for message in messages %}
    #     <li>{{ message }}</li>
    #   {% endfor %}
    #   </ul>
    # {% endif %}
    # {% endwith %}
    # <h1>Загрузить новый файл</h1>
    # <form method=post enctype=multipart/form-data>
    #   <input type=file name=file>
    #   <input type=submit value=Upload>
    # </form>
    # </html>
    # '''


if __name__ == "__main__":
    app.run(debug=True)
    # app.run()
