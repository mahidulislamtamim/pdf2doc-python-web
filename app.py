from flask import Flask, render_template, request, send_from_directory
from pdf2docx import Converter
import os

app = Flask(__name__)

def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")

# Define folders 
script_dir = os.path.dirname(os.path.abspath(__file__))
upload_folder = os.path.join(script_dir, "upload")
output_folder = os.path.join(script_dir, "converted")

# Clear folder
clear_folder(upload_folder)
clear_folder(output_folder)

# Create folder
os.makedirs(upload_folder, exist_ok=True)
os.makedirs(output_folder, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf_input"]
        if file and file.filename.endswith(".pdf"):
            file_name = file.filename
            pdf_path = os.path.join(upload_folder, file_name)
            docx_name = file_name.replace(".pdf", ".docx")
            docx_path = os.path.join(output_folder, docx_name)

            #save pdf
            file.save(pdf_path)

            #Convert pdf to docx
            converter = Converter(pdf_path)
            converter.convert(docx_path)
            converter.close()

            return send_from_directory(output_folder, docx_name, as_attachment = True)
        
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)