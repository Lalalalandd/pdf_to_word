import asyncio

async def convert_pdf_to_doc(pdf_file):
    command = [
        "libreoffice", "--headless", "--infilter=writer_pdf_import",
        "--convert-to", "doc:MS Word 97",
        pdf_file
    ]
    
    process = await asyncio.create_subprocess_exec(
        *command,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )

    await process.wait()

    if process.returncode == 0:
        print(f"✅ Conversion successful: {pdf_file} → .doc")
    else:
        stderr = await process.stderr.read()
        print(f"❌ Error: {stderr.decode()}")

pdf_file = "your_document_path"
asyncio.run(convert_pdf_to_doc(pdf_file))
