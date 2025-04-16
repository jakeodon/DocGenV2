import azure.functions as func
import logging
from io import BytesIO

from docxtpl import DocxTemplate  # Replacing python-docx with docxtpl for templating
from pypdf import PdfWriter
import json
import os
import requests

from dotenv import load_dotenv
load_dotenv()

# Environment variables for Azure credentials
tenant_id = os.getenv("AZURE_TENANT_ID")
client_id = os.getenv("AZURE_CLIENT_ID")
client_secret = os.getenv("AZURE_CLIENT_SECRET")
account_id = os.getenv("AZURE_ACCOUNT_ID")

app = func.FunctionApp()

@app.function_name(name="ReplacePlaceholders")
@app.route(route="replace", methods=["POST"])
def replace_placeholders_function(req: func.HttpRequest) -> func.HttpResponse:
    """
    Parses a Word document and replaces {{key}} placeholders with the values from an input JSON file.
    """
    logging.info("Processing request to replace placeholders in Word document.")

    try:
        # Parse the incoming request
        # Get the Word document, JSON data, and output file name from the request body
        word_file = req.files.get("word_file")
        json_file = req.files.get("json_file")
        output_file_name = req.params.get("output_file_name")  # Get the output file name from query params

        if not word_file or not json_file:
            return func.HttpResponse(
                "Please provide both a Word document and a JSON file.",
                status_code=400,
            )

        if not output_file_name:
            output_file_name = "modified.docx"  # Default file name if none is provided

        # Load the Word template using docxtpl
        doc = DocxTemplate(word_file.stream)

        # Load JSON data from the uploaded file
        json_data = json.load(json_file.stream)

        # Render the template with the JSON data
        doc.render(json_data)

        # Save the modified document to an in-memory stream
        output_stream = BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)  # Reset the stream position to the beginning

        # Return the modified Word document as a response
        return func.HttpResponse(
            output_stream.read(),
            status_code=200,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={output_file_name}"},
        )

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return func.HttpResponse(
            f"An error occurred while processing the request: {str(e)}",
            status_code=500,
        )


def get_access_token():
    """
    Get an access token from Azure AD for Microsoft Graph API.
    """
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(url, data=data)
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        logging.error(f"Failed to get token: {response.status_code} - {response.text}")
        raise Exception("Failed to get access token")


def upload_to_graph(file_name: str, file_content: bytes, access_token: str):
    """
    Upload a Word document to Microsoft Graph API.
    """
    upload_url = f"https://graph.microsoft.com/v1.0/users/{account_id}/drive/root:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    response = requests.put(upload_url, headers=headers, data=file_content)
    if response.status_code in [200, 201]:
        logging.info(f"File '{file_name}' uploaded successfully.")
    else:
        logging.error(f"Error uploading file: {response.status_code} - {response.text}")
        raise Exception("Failed to upload file to Microsoft Graph")


def convert_to_pdf(file_name: str, access_token: str) -> bytes:
    """
    Convert a Word document to PDF using Microsoft Graph API.
    """
    convert_url = f"https://graph.microsoft.com/v1.0/users/{account_id}/drive/root:/{file_name}:/content?format=pdf"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(convert_url, headers=headers)
    if response.status_code == 200:
        logging.info(f"File '{file_name}' successfully converted to PDF.")
        return response.content
    else:
        logging.error(f"Error converting file to PDF: {response.status_code} - {response.text}")
        raise Exception("Failed to convert file to PDF")


@app.function_name(name="Word2Pdf")
@app.route(route="word-to-pdf", methods=["post"])
def word_to_pdf_function(req: func.HttpRequest) -> func.HttpResponse:
    """
    Azure Function to convert a Word document to PDF using Microsoft Graph API.
    """
    logging.info("Word2Pdf function processing a request.")

    try:
        # Get the Word document from the request
        word_file = req.files.get('word_file')
        if not word_file:
            return func.HttpResponse(
                "Please upload a Word document as 'word_file' in the request.",
                status_code=400
            )

        # Extract the file name and content
        original_file_name = word_file.filename
        if not original_file_name.endswith(".docx"):
            return func.HttpResponse(
                "Only .docx files are supported.",
                status_code=400
            )

        word_file_content = word_file.read()

        # Get an access token
        access_token = get_access_token()

        # Upload the Word document to Microsoft Graph
        upload_to_graph(original_file_name, word_file_content, access_token)

        # Convert the Word document to PDF
        pdf_content = convert_to_pdf(original_file_name, access_token)

        # Generate the output PDF file name
        pdf_file_name = original_file_name.rsplit(".", 1)[0] + ".pdf"

        # Return the PDF file as a response
        return func.HttpResponse(
            body=pdf_content,
            status_code=200,
            mimetype="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={pdf_file_name}"
            }
        )

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return func.HttpResponse(
            "An error occurred while processing the request. Please ensure the file is valid.",
            status_code=500
        )


from pypdf import PdfWriter  # Import PdfWriter instead of PdfMerger

@app.function_name(name="CombinePages")
@app.route(route="combine-pages", methods=["post"])
def combine_pages_function(req: func.HttpRequest) -> func.HttpResponse:
    """
    Azure Function to combine multiple PDF files into a single PDF.
    """
    logging.info("CombinePages function processing a request.")

    try:
        # Get the uploaded PDF files from the request
        pdf_files = req.files.getlist('pdf_files')
        if not pdf_files or len(pdf_files) < 2:
            return func.HttpResponse(
                "Please upload at least two PDF files as 'pdf_files' in the request.",
                status_code=400
            )

        # Initialize the PDF writer
        writer = PdfWriter()

        # Iterate through the uploaded files and add them to the writer
        for pdf_file in pdf_files:
            pdf_file_bytes = BytesIO(pdf_file.read())
            writer.append(pdf_file_bytes)

        # Save the merged PDF to a BytesIO stream
        combined_pdf = BytesIO()
        writer.write(combined_pdf)
        combined_pdf.seek(0)  # Reset the stream position to the beginning

        # Return the combined PDF as a response
        return func.HttpResponse(
            body=combined_pdf.getvalue(),
            status_code=200,
            mimetype="application/pdf",
            headers={
                "Content-Disposition": "attachment; filename=combined_document.pdf"
            }
        )

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return func.HttpResponse(
            "An error occurred while processing the request. Please ensure the files are valid.",
            status_code=500
        )
    

    