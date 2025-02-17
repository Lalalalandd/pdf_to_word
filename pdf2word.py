import subprocess

pdf_file = "your_document_path"

command = [
    "libreoffice", "--headless", "--infilter=writer_pdf_import",
    "--convert-to", "doc:MS Word 97",
    pdf_file
]

subprocess.run(command, check=True)
