import os
from flask import Flask, request, send_file, redirect, flash
from werkzeug.utils import secure_filename
import xport.v56
from flask_httpauth import HTTPBasicAuth
from werkzeug.security import generate_password_hash, check_password_hash

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"xpt"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # Limit upload size to 100MB

auth = HTTPBasicAuth()

users = {"user": generate_password_hash("password")}


@auth.verify_password
def verify_password(username, password):
    if username in users and check_password_hash(users.get(username), password):
        return username


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
@auth.login_required
def upload_form():
    return """
    <!doctype html>
    <title>Upload XPT File</title>
    <h1>Upload XPT File</h1>
    <form method="post" enctype="multipart/form-data" action="/convert">
      <input type="file" name="file">
      <input type="submit" value="Convert to Excel">
    </form>
    """


@app.route("/convert", methods=["POST"])
@auth.login_required
def convert_file():
    if "file" not in request.files:
        flash("No file part")
        return redirect(request.url)
    file = request.files["file"]
    if file.filename == "":
        flash("No selected file")
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)

        # Convert file to Excel
        target = filename.rsplit(".", 1)[0]
        with open(file_path, "rb") as f:
            library = xport.v56.load(f)
        member = library.get("DRXIFF")
        header = member.contents.get("Label").to_dict().values()
        output_filename = f"{target}.xlsx"
        member.to_excel(output_filename, header=header)

        return send_file(output_filename, as_attachment=True)
    else:
        flash("Allowed file types are xpt")
        return redirect(request.url)


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True)
