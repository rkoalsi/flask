from flask import Flask, request, jsonify, send_file
from helpers import process_upload, validate_file
import logging, threading


app = Flask(__name__)
app.logger.setLevel(logging.DEBUG)


@app.route("/", methods=["GET"])
def index():
    return "Flask App is running successfully"


@app.route("/hello", methods=["GET"])
def hello_world():
    return {"data": "Hello, World!"}


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return {"error": "No file part"}, 400
    email = str(request.form.get("email"))
    if not email:
        return jsonify({"error": "Email is required"}), 400
    file = request.files["file"]
    if file.filename == "":
        return {"error": "No selected file"}, 400
    r = validate_file(file)
    status = r.get("status")
    message = r.get("message")
    if status == "error":
        return {"message": f"Error in file uploaded, {message}"}, 400
    try:
        # Start processing in a separate thread
        threading.Thread(target=process_upload, args=(file, email)).start()

        # Return a response immediately
        return {
            "message": f"Processing started.\nAn email will be sent to {email} once the task is completed."
        }

    except Exception as e:
        return {"error": f"Error processing file: {e}"}, 500

    except Exception as e:
        return {"error": f"Error saving file: {e}"}, 500


@app.route("/download", methods=["GET"])
def download():
    name = "Template.xlsx"
    try:
        return send_file(name, as_attachment=True, download_name=name)
    except Exception as e:
        return str(e)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
