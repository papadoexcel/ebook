from flask import Flask, request, render_template_string, send_file
import mammoth
from ebooklib import epub
from xhtml2pdf import pisa
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

HTML_FORM = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Conversor de Word para eBook</title>
    <style>
        body {
            background: #f4f4f4;
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }
        .container {
            background: white;
            padding: 30px 40px;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 500px;
        }
        h2 {
            text-align: center;
            color: #333;
            margin-bottom: 20px;
        }
        label {
            font-weight: bold;
            display: block;
            margin-top: 15px;
        }
        input[type="text"], select, input[type="file"] {
            width: 100%;
            padding: 10px;
            margin-top: 5px;
            border-radius: 5px;
            border: 1px solid #ccc;
        }
        input[type="submit"] {
            background: #0066cc;
            color: white;
            padding: 12px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            margin-top: 20px;
            width: 100%;
            font-size: 16px;
        }
        input[type="submit"]:hover {
            background: #004d99;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Conversor de Word para eBook</h2>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <label for="title">Título do eBook:</label>
            <input type="text" name="title" required>

            <label for="author">Autor:</label>
            <input type="text" name="author" required>

            <label for="format">Formato de saída:</label>
            <select name="format" required>
                <option value="epub">EPUB</option>
                <option value="pdf">PDF</option>
            </select>

            <label for="file">Arquivo Word (.docx):</label>
            <input type="file" name="file" accept=".docx" required>

            <input type="submit" value="Converter">
        </form>
    </div>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_FORM)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    title = request.form['title']
    author = request.form['author']
    output_format = request.form['format']

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    with open(filepath, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value

    if output_format == 'epub':
        book = epub.EpubBook()
        book.set_title(title)
        book.add_author(author)

        chapter = epub.EpubHtml(title="Conteúdo", file_name="chap1.xhtml", content=html, lang="pt")
        book.add_item(chapter)
        book.toc = (epub.Link('chap1.xhtml', 'Conteúdo', 'chap1'),)
        book.spine = ['nav', chapter]
        book.add_item(epub.EpubNav())
        book.add_item(epub.EpubNcx())

        epub_filename = filename.replace(".docx", ".epub")
        epub_path = os.path.join(OUTPUT_FOLDER, epub_filename)
        epub.write_epub(epub_path, book)
        return send_file(epub_path, as_attachment=True)

    elif output_format == 'pdf':
        pdf_filename = filename.replace(".docx", ".pdf")
        pdf_path = os.path.join(OUTPUT_FOLDER, pdf_filename)

        with open(pdf_path, "wb") as pdf_file:
            pisa_status = pisa.CreatePDF(html, dest=pdf_file)

        if pisa_status.err:
            return "Erro ao gerar PDF", 500

        return send_file(pdf_path, as_attachment=True)

    else:
        return "Formato inválido", 400


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

