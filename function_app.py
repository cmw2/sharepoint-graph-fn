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

@app.route(route="sharepoint_docs_list", methods=["GET"])
def sharepoint_docs_list(req: func.HttpRequest) -> func.HttpResponse:
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
        
        # Get all documents
        docs = client.list_documents()
        
        # Return the documents as JSON
        import json
        return func.HttpResponse(
            json.dumps({"documents": docs}),
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
