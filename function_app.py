import azure.functions as func
import logging
import os
import traceback
from sharepoint_graph import SharePointGraphClient

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="httpget", methods=["GET"])
def http_get(req: func.HttpRequest) -> func.HttpResponse:
    name = req.params.get("name", "World")

    logging.info(f"Processing GET request. Name: {name}")

    return func.HttpResponse(f"Hello, {name}!")

@app.route(route="httppost", methods=["POST"])
def http_post(req: func.HttpRequest) -> func.HttpResponse:
    try:
        req_body = req.get_json()
        name = req_body.get('name')
        age = req_body.get('age')
        
        logging.info(f"Processing POST request. Name: {name}")

        if name and isinstance(name, str) and age and isinstance(age, int):
            return func.HttpResponse(f"Hello, {name}! You are {age} years old!")
        else:
            return func.HttpResponse(
                "Please provide both 'name' and 'age' in the request body.",
                status_code=400
            )
    except ValueError:
        return func.HttpResponse(
            "Invalid JSON in request body",
            status_code=400
        )

# @app.schedule(schedule="0 */4 * * *", arg_name="timer", run_on_startup=False) # Run every 4 hours
# def sharepoint_document_processor(timer: func.TimerRequest) -> None:
#     """
#     Azure Function to process SharePoint documents.
#     This function runs on a schedule and logs information about the documents.
#     """
#     logging.info("Sharepoint Document Processor function triggered")
    
#     # Check environment variable configuration
#     sp_tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
#     site_name = os.environ.get("SHAREPOINT_SITE_NAME")
#     document_library = os.environ.get("SHAREPOINT_DOCUMENT_LIBRARY", "Documents")
    
#     # Check for required configuration
#     if not sp_tenant_id:
#         logging.error("SHAREPOINT_TENANT_ID environment variable must be set")
#         return
        
#     if not site_name:
#         logging.error("SHAREPOINT_SITE_NAME environment variable must be set")
#         return
        
#     try:
#         # Initialize client
#         client = SharePointGraphClient(
#             sp_tenant_id=sp_tenant_id,
#             site_name=site_name,
#             document_library=document_library
#         )
        
#         # Process all documents
#         processed_docs = client.process_all_documents()
        
#         logging.info(f"Successfully processed {len(processed_docs)} files")
        
#     except Exception as e:
#         logging.error(f"An error occurred: {str(e)}")
#         logging.error(traceback.format_exc())

@app.route(route="sharepoint", methods=["GET"])
def list_sharepoint_docs(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP Trigger function to list SharePoint documents.
    This function can be called anytime to get a list of documents.
    """
    logging.info("SharePoint documents listing function triggered")
    
    try:
        # Check environment variable configuration
        sp_tenant_id = os.environ.get("SHAREPOINT_TENANT_ID")
        site_name = os.environ.get("SHAREPOINT_SITE_NAME")
        document_library = os.environ.get("SHAREPOINT_DOCUMENT_LIBRARY", "Documents")
        
        # Check for required configuration
        if not sp_tenant_id:
            return func.HttpResponse(
                "SHAREPOINT_TENANT_ID environment variable must be set",
                status_code=400
            )
            
        if not site_name:
            return func.HttpResponse(
                "SHAREPOINT_SITE_NAME environment variable must be set",
                status_code=400
            )
        
        # Initialize client
        client = SharePointGraphClient(
            sp_tenant_id=sp_tenant_id,
            site_name=site_name,
            document_library=document_library
        )
        
        # Process all documents
        processed_docs = client.process_all_documents()
        
        # Return the documents as JSON
        import json
        return func.HttpResponse(
            json.dumps({"documents": processed_docs}),
            mimetype="application/json"
        )
        
    except Exception as e:
        error_message = f"An error occurred: {str(e)}"
        logging.error(error_message)
        logging.error(traceback.format_exc())
        return func.HttpResponse(
            error_message,
            status_code=500
        )
