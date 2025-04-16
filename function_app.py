import azure.functions as func
import logging
from io import BytesIO

from docxtpl import DocxTemplate  # Replacing python-docx with docxtpl for templating
from pypdf import PdfWriter
import json
import os
import requests

from dotenv import load_dotenv
import uuid
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
    Parses a Word document from a signed URL, replaces {{key}} placeholders with JSON data,
    and saves the output as a PDF to another signed URL.
    """
    logging.info("Processing request to replace placeholders in Word document.")

    try:
        # Parse the incoming request
        request_data = req.get_json()
        input_uri = request_data.get("input", {}).get("uri")
        output_uri = request_data.get("output", {}).get("uri")
        json_data_for_merge = request_data.get("params", {}).get("jsonDataForMerge", {})
        output_format = request_data.get("params", {}).get("outputFormat", "pdf")

        if not input_uri or not output_uri or not json_data_for_merge:
            logging.error("Missing required fields: 'input.uri', 'output.uri', or 'params.jsonDataForMerge'.")
            return func.HttpResponse(
                "Please provide 'input.uri', 'output.uri', and 'params.jsonDataForMerge' in the request.",
                status_code=400,
            )

        # Fetch the Word document from the input URI
        logging.info(f"Fetching Word document from input URI: {input_uri}")
        response = requests.get(input_uri)
        if response.status_code != 200:
            logging.error(f"Failed to fetch the Word document. Status code: {response.status_code}")
            return func.HttpResponse(
                f"Failed to fetch the Word document from the input URI. Status code: {response.status_code}",
                status_code=400,
            )

        logging.info(f"Successfully fetched Word document. Content size: {len(response.content)} bytes.")
        word_file_content = BytesIO(response.content)

        # Load the Word template using docxtpl
        logging.info("Loading Word template using docxtpl.")
        doc = DocxTemplate(word_file_content)

        # Render the template with the JSON data
        logging.info("Rendering the Word template with JSON data.")
        doc.render(json_data_for_merge)

        # Save the modified document to an in-memory stream
        output_stream = BytesIO()
        if output_format.lower() == "pdf":
            # Convert to PDF if requested
            logging.info("Converting the Word document to PDF.")
            temp_docx_stream = BytesIO()
            doc.save(temp_docx_stream)
            temp_docx_stream.seek(0)

            logging.info(f"Temporary DOCX stream size: {temp_docx_stream.getbuffer().nbytes} bytes.")

            word_file_content = temp_docx_stream
            original_file_name = f"{uuid.uuid4()}.docx"
        # Get an access token
            access_token = get_access_token()
        # Upload the Word document to Microsoft Graph
            upload_to_graph(original_file_name, word_file_content, access_token)

        # Convert the Word document to PDF
            pdf_content = convert_to_pdf(original_file_name, access_token)

        # Generate the output PDF file name
            pdf_file_name = original_file_name.rsplit(".", 1)[0] + ".pdf"

            logging.info(f"PDF conversion completed. Output stream size: {output_stream.getbuffer().nbytes} bytes.")
        else:
            # Save as DOCX
            logging.info("Saving the modified document as DOCX.")
            doc.save(output_stream)
            logging.info(f"DOCX saved. Output stream size: {output_stream.getbuffer().nbytes} bytes.")

        output_stream.seek(0)  # Reset the stream position to the beginning

        # Save the output to the output URI
        logging.info(f"Uploading the output file to the output URI: {output_uri}")
        put_response = requests.put(output_uri, data=pdf_content, headers={
            "Content-Type": "application/pdf" if output_format.lower() == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        })
        if put_response.status_code not in [200, 201]:
            logging.error(f"Failed to save the output file. Status code: {put_response.status_code}")
            return func.HttpResponse(
                f"Failed to save the output file to the output URI. Status code: {put_response.status_code}",
                status_code=400,
            )

        logging.info("File processed and saved successfully.")
        return func.HttpResponse(
            "File processed and saved successfully.",
            status_code=201,
        )

    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True)
        return func.HttpResponse(
            f"An error occurred while processing the request: {str(e)}",
            status_code=500,
        )

@app.function_name(name="CombinePages")
@app.route(route="combine-pages", methods=["POST"])
def combine_pages_function(req: func.HttpRequest) -> func.HttpResponse:
    """
    Azure Function to combine multiple PDF files from signed URLs into a single PDF.
    """
    logging.info("CombinePages function processing a request.")
    logging.info("Request body: %s", req.get_body())


    try:
        # Parse the incoming request
        request_data = req.get_json()
        input_uris = request_data.get("inputs",[])
        output_uri = request_data.get("output", {}).get("uri")

        if not input_uris or len(input_uris) < 2 or not output_uri:
            logging.error("Missing required fields: 'input.uris' (at least two) or 'output.uri'.")
            return func.HttpResponse(
                "Please provide at least two 'input.uris' and an 'output.uri' in the request.",
                status_code=400,
            )

        # Initialize the PDF writer
        writer = PdfWriter()

        # Fetch and append each PDF from the input URIs
        for uri in input_uris:
            logging.info(f"Fetching PDF from URI: {uri}")
            response = requests.get(uri.get("input", {}).get("uri"))
            if response.status_code != 200:
                logging.error(f"Failed to fetch PDF from URI: {uri}. Status code: {response.status_code}")
                return func.HttpResponse(
                    f"Failed to fetch PDF from URI: {uri}. Status code: {response.status_code}",
                    status_code=400,
                )
            pdf_file_bytes = BytesIO(response.content)
            writer.append(pdf_file_bytes)

        # Save the combined PDF to an in-memory stream
        combined_pdf = BytesIO()
        writer.write(combined_pdf)
        combined_pdf.seek(0)  # Reset the stream position to the beginning

        # Save the combined PDF to the output URI
        logging.info(f"Uploading the combined PDF to the output URI: {output_uri}")
        put_response = requests.put(output_uri, data=combined_pdf.getvalue(), headers={
            "Content-Type": "application/pdf"
        })

        if put_response.status_code not in [200, 201]:
            logging.error(f"Failed to save the combined PDF. Status code: {put_response.status_code}")
            return func.HttpResponse(
                f"Failed to save the combined PDF to the output URI. Status code: {put_response.status_code}",
                status_code=400,
            )

        logging.info("Combined PDF saved successfully.")
        return func.HttpResponse(
            "Combined PDF saved successfully.",
            status_code=201,
        )

    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True)
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




# @app.function_name(name="Word2Pdf")
# @app.route(route="word-to-pdf", methods=["post"])
# def word_to_pdf_function(req: func.HttpRequest) -> func.HttpResponse:
#     """
#     Azure Function to convert a Word document to PDF using Microsoft Graph API.
#     """
#     logging.info("Word2Pdf function processing a request.")

#     try:
#         # Get the Word document from the request
#         word_file = req.files.get('word_file')
#         if not word_file:
#             return func.HttpResponse(
#                 "Please upload a Word document as 'word_file' in the request.",
#                 status_code=400
#             )

#         # Extract the file name and content
#         original_file_name = word_file.filename
#         if not original_file_name.endswith(".docx"):
#             return func.HttpResponse(
#                 "Only .docx files are supported.",
#                 status_code=400
#             )

#         word_file_content = word_file.read()

#         # Get an access token
#         access_token = get_access_token()

#         # Upload the Word document to Microsoft Graph
#         upload_to_graph(original_file_name, word_file_content, access_token)

#         # Convert the Word document to PDF
#         pdf_content = convert_to_pdf(original_file_name, access_token)

#         # Generate the output PDF file name
#         pdf_file_name = original_file_name.rsplit(".", 1)[0] + ".pdf"

#         # Return the PDF file as a response
#         return func.HttpResponse(
#             body=pdf_content,
#             status_code=200,
#             mimetype="application/pdf",
#             headers={
#                 "Content-Disposition": f"attachment; filename={pdf_file_name}"
#             }
#         )

#     except Exception as e:
#         logging.error(f"An error occurred: {e}")
#         return func.HttpResponse(
#             "An error occurred while processing the request. Please ensure the file is valid.",
#             status_code=500
#         )