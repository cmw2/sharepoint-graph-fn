"""
SharePoint Graph API Client Library

This module provides functionality to interact with SharePoint document libraries via Microsoft Graph API.
"""
import os
import time
import json
import logging
import traceback
import requests
from typing import List, Dict, Any, Optional
from azure.identity import DefaultAzureCredential

# Configure logging
logger = logging.getLogger('sharepoint-graph')

class SharePointGraphClient:
    """Client for accessing SharePoint documents via Microsoft Graph API"""

    # Microsoft Graph API endpoint
    GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
    
    # Required scopes for accessing SharePoint documents
    GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

    def __init__(
        self,
        sp_tenant_id: str = None,
        site_name: str = None,
        document_library: str = None
    ):
        """
        Initialize the SharePoint Graph client.
        
        Args:
            sp_tenant_id: Azure tenant ID (required)
            site_name: SharePoint site name (required)
            document_library: Document library name (required)
        """
        self.sp_tenant_id = sp_tenant_id or os.environ.get("SHAREPOINT_TENANT_ID")
        if not self.sp_tenant_id:
            raise ValueError("Tenant ID must be provided or set as SHAREPOINT_TENANT_ID environment variable")
        
        self.site_name = site_name or os.environ.get("SHAREPOINT_SITE_NAME")
        if not self.site_name:
            raise ValueError("SharePoint site name must be provided or set as SHAREPOINT_SITE_NAME environment variable")
        
        self.document_library = document_library or os.environ.get("SHAREPOINT_DOCUMENT_LIBRARY", "Documents")
        
        # Initialize credential
        self.credential = DefaultAzureCredential()
        self.token = None
        self.token_expires_at = 0

    def _ensure_token(self) -> None:
        """Ensure we have a valid access token"""
        current_time = time.time()
        
        # If token is expired or will expire in next 5 minutes, refresh it
        if self.token is None or current_time >= (self.token_expires_at - 300):
            logger.info("Getting new access token")
            token = self.credential.get_token(*self.GRAPH_SCOPES)
            self.token = token.token
            self.token_expires_at = token.expires_on
            logger.info(f"Token acquired, expires at: {self.token_expires_at}")

    def _make_request(
        self, 
        method: str, 
        endpoint: str, 
        params: Dict[str, Any] = None, 
        headers: Dict[str, Any] = None, 
        json_data: Dict[str, Any] = None
    ) -> requests.Response:
        """
        Make an authenticated request to the Microsoft Graph API.
        
        Args:
            method: HTTP method (GET, POST, etc.)
            endpoint: API endpoint to call
            params: Query parameters
            headers: Request headers
            json_data: JSON data for POST/PUT requests

        Returns:
            Response object
        """
        self._ensure_token()
        
        url = f"{self.GRAPH_API_ENDPOINT}{endpoint}"
        
        if headers is None:
            headers = {}
            
        headers.update({
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        })
        
        # Implement retry logic with exponential backoff
        max_retries = 3
        retry_delay = 1  # Start with 1 second delay
        
        for attempt in range(max_retries):
            try:
                #logger.info(f"Making {method} request to {url} with params: {params} and headers: {headers}")
                logger.info(f"Making {method} request to {url} with params: {params}")
                response = requests.request(
                    method=method,
                    url=url,
                    params=params,
                    headers=headers,
                    json=json_data,
                    timeout=30  # Set a reasonable timeout
                )
                
                # Check if request was successful
                response.raise_for_status()
                return response
                
            except (requests.RequestException, ConnectionError) as e:
                if attempt < max_retries - 1:
                    logger.warning(f"Request failed (attempt {attempt+1}/{max_retries}): {str(e)}. Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # Exponential backoff
                else:
                    logger.error(f"Request failed after {max_retries} attempts: {str(e)}")
                    raise
        
        # This shouldn't be reached due to the raise in the exception handler, but added for completeness
        return None

    def get_site_id(self) -> str:
        """
        Get the SharePoint site ID based on the site name.
        
        Returns:
            The site ID
            
        Raises:
            ValueError: If site ID cannot be retrieved, with details about the response
        """
        logger.info(f"Getting site ID for: {self.site_name}")
        
        # Format the request for site lookup
        endpoint = f"/sites/{self.sp_tenant_id}.sharepoint.com:/sites/{self.site_name}"
        
        response = self._make_request("GET", endpoint)
        site_data = response.json()
        
        site_id = site_data.get("id")
        if not site_id:
            # Log the entire response for debugging
            logger.error(f"Failed to retrieve site ID. Response: {json.dumps(site_data, indent=2)}")
            error_msg = site_data.get("error", {}).get("message", "No specific error message provided")
            raise ValueError(f"Could not retrieve site ID for {self.site_name}. Error: {error_msg}. Response: {site_data}")
            
        logger.info(f"Retrieved site ID: {site_id}")
        return site_id

    def get_drive_id(self, site_id: str) -> str:
        """
        Get the drive ID for the document library.
        
        Args:
            site_id: The SharePoint site ID
            
        Returns:
            The drive ID for the document library
            
        Raises:
            ValueError: If drive ID cannot be retrieved, with details about available drives
        """
        logger.info(f"Getting drive ID for document library: {self.document_library}")
        
        # First try to get the drive by name
        endpoint = f"/sites/{site_id}/drives"
        
        response = self._make_request("GET", endpoint)
        response_data = response.json()
        drives = response_data.get("value", [])
        
        # Log the available drive names for debugging
        available_drives = [drive.get("name") for drive in drives]
        logger.info(f"Available drives: {available_drives}")
        
        for drive in drives:
            if drive.get("name") == self.document_library:
                drive_id = drive.get("id")
                logger.info(f"Found drive ID for {self.document_library}: {drive_id}")
                return drive_id
                
        # If not found by name, try to get the default document library
        logger.warning(f"Document library '{self.document_library}' not found. Trying to get default document library.")
        
        for drive in drives:
            if drive.get("name") == "Documents":
                drive_id = drive.get("id")
                logger.info(f"Found default drive ID: {drive_id}")
                return drive_id
        
        # If we get here, no suitable drive was found
        logger.error(f"Failed to find drive. Response data: {json.dumps(response_data, indent=2)}")
        available_drives_str = ", ".join(available_drives) if available_drives else "No drives found"
        raise ValueError(f"Could not find drive for document library: {self.document_library}. Available drives: {available_drives_str}")

    def list_documents(self, site_id: str, drive_id: str, folder_path: str = "") -> List[Dict[str, Any]]:
        """
        List documents in the specified folder.
        
        Args:
            site_id: The SharePoint site ID
            drive_id: The drive ID
            folder_path: Optional path to a subfolder
            
        Returns:
            List of document metadata
            
        Raises:
            ValueError: If folder path is invalid or permissions issue
            requests.RequestException: If API request fails with detailed error information
        """
        logger.info(f"Listing documents in folder: '{folder_path or 'root'}'")
        
        # Format the endpoint
        endpoint = f"/sites/{site_id}/drives/{drive_id}/root"
        if folder_path:
            endpoint = f"{endpoint}:/{folder_path}:"
        endpoint = f"{endpoint}/children"
        
        params = {
            "$select": "id,name,size,webUrl,file,folder",
            "$expand": "thumbnails",
            "$top": 1000  # Adjust as needed
        }
        
        try:
            response = self._make_request("GET", endpoint, params=params)
            response_data = response.json()
            
            # Check if there's an error in the response
            if "error" in response_data:
                error = response_data["error"]
                error_message = error.get("message", "Unknown error")
                logger.error(f"Error listing documents: {error_message}")
                raise ValueError(f"Failed to list documents: {error_message}")
            
            items = response_data.get("value", [])
            logger.info(f"Found {len(items)} items in folder '{folder_path or 'root'}'")
        except requests.RequestException as e:
            logger.error(f"Failed to list documents in folder '{folder_path}': {str(e)}", exc_info=True)
            raise
        
        documents = []
        folders = []
        
        # Process the returned items
        for item in items:
            if "folder" in item:
                folder_name = item.get("name", "")
                folder_path_new = f"{folder_path}/{folder_name}" if folder_path else folder_name
                folders.append({"name": folder_name, "path": folder_path_new})
            elif "file" in item:
                documents.append({
                    "id": item.get("id", ""),
                    "name": item.get("name", ""),
                    "path": folder_path,
                    "size": item.get("size", 0),
                    "web_url": item.get("webUrl", "")
                })
                
        # Recursively get documents from subfolders
        for folder in folders:
            subfolder_docs = self.list_documents(site_id, drive_id, folder["path"])
            documents.extend(subfolder_docs)
            
        return documents

    def log_document(self, document: Dict[str, Any]) -> None:
        """
        Log document information (instead of downloading).
        
        Args:
            document: Document metadata
        """
        doc_id = document.get("id")
        doc_name = document.get("name")
        doc_path = document.get("path", "")
        doc_size = document.get("size", 0)
        doc_url = document.get("web_url", "")
        
        # Validate document metadata
        if not doc_id or not doc_name:
            missing_fields = []
            if not doc_id: 
                missing_fields.append("id")
            if not doc_name: 
                missing_fields.append("name")
            logger.warning(f"Incomplete document metadata. Missing fields: {', '.join(missing_fields)}. Document data: {document}")
            return
        
        full_path = f"{doc_path}/{doc_name}" if doc_path else doc_name
        logger.info(f"Found document: {full_path} (ID: {doc_id}, Size: {doc_size} bytes, URL: {doc_url})")

    def process_all_documents(self) -> List[Dict[str, Any]]:
        """
        Process all documents from the document library.
        
        Returns:
            List of document information
        """
        # Get IDs needed for the operations
        site_id = self.get_site_id()
        drive_id = self.get_drive_id(site_id)
        
        # List all documents
        documents = self.list_documents(site_id, drive_id)
        
        processed_documents = []
        
        # Process each document
        for document in documents:
            try:
                self.log_document(document)
                processed_documents.append(document)
            except Exception as e:
                logger.error(f"Error processing document {document.get('name', 'Unknown')}: {str(e)}")
                
        logger.info(f"Processed {len(processed_documents)} documents")
        return processed_documents
