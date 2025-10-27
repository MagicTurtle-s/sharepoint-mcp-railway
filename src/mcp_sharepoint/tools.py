import base64, os
from functools import wraps
from typing import Optional, Dict, Any
from .common import logger, mcp, SHP_DOC_LIBRARY, sp_context, get_sp_context_for_site
from .resources import list_folders, list_documents, get_document_content, get_folder_tree, download_document
from .graph_api import list_sharepoint_sites as _list_sharepoint_sites

# Helper functions to reduce code duplication
def _get_path(folder: str = "", file: Optional[str] = None) -> str:
    """Construct SharePoint path from components"""
    path = f"{SHP_DOC_LIBRARY}/{folder}".rstrip('/')
    return f"{path}/{file}" if file else path

def _handle_sp_operation(func):
    """Decorator for SharePoint operations with error handling"""
    @wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            return {"success": False, "message": f"Operation failed: {str(e)}"}
    return wrapper

def _file_success_response(file_obj, message: str) -> Dict[str, Any]:
    """Standard success response for file operations"""
    return {
        "success": True,
        "message": message,
        "file": {"name": file_obj.name, "url": file_obj.serverRelativeUrl}
    }

# ===== MULTI-SITE TOOLS =====

@mcp.tool(
    name="List_SharePoint_Sites",
    description="Discover all accessible SharePoint sites in the tenant. Use this first to find site URLs for other operations."
)
async def list_sharepoint_sites_tool(search: Optional[str] = None):
    """
    List all SharePoint sites accessible to the application.

    This tool uses Microsoft Graph API to discover sites across the tenant.
    Use the returned site_url values in other tools to access specific sites.

    Args:
        search: Optional search term to filter sites by name or description

    Returns:
        List of sites with name, site_url, site_id, and description
    """
    try:
        sites = await _list_sharepoint_sites(search)
        return {
            "success": True,
            "count": len(sites),
            "sites": sites
        }
    except Exception as e:
        logger.error(f"Error listing SharePoint sites: {str(e)}")
        return {
            "success": False,
            "message": f"Failed to list sites: {str(e)}"
        }


# ===== DOCUMENT & FOLDER TOOLS (Multi-Site Ready) =====

@mcp.tool(
    name="List_SharePoint_Folders",
    description="List folders in a SharePoint site directory. Requires site_url from List_SharePoint_Sites."
)
async def list_folders_tool(site_url: str, parent_folder: Optional[str] = None, doc_library: Optional[str] = None):
    """
    List folders in the specified SharePoint directory or root if not specified.

    Args:
        site_url: Full SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)
        parent_folder: Relative folder path within doc library (optional, defaults to root)
        doc_library: Document library name (optional, defaults to "Shared Documents")
    """
    try:
        ctx = get_sp_context_for_site(site_url)
        return list_folders(parent_folder, ctx)
    except Exception as e:
        logger.error(f"Error listing folders: {str(e)}")
        return {"error": str(e)}

@mcp.tool(
    name="List_SharePoint_Documents",
    description="List all documents in a SharePoint folder. Requires site_url from List_SharePoint_Sites."
)
async def list_documents_tool(site_url: str, folder_name: str, doc_library: Optional[str] = None):
    """
    List all documents in a specified SharePoint folder.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Relative folder path within doc library
        doc_library: Document library name (optional, defaults to "Shared Documents")
    """
    try:
        ctx = get_sp_context_for_site(site_url)
        return list_documents(folder_name, ctx)
    except Exception as e:
        logger.error(f"Error listing documents: {str(e)}")
        return {"error": str(e)}

@mcp.tool(
    name="Get_SharePoint_Tree",
    description="Get a recursive tree view of a SharePoint folder structure. Requires site_url."
)
async def get_sharepoint_tree_tool(site_url: str, parent_folder: Optional[str] = None, doc_library: Optional[str] = None):
    """
    Get a recursive tree view of a SharePoint folder.

    Args:
        site_url: Full SharePoint site URL
        parent_folder: Relative folder path (optional, defaults to root)
        doc_library: Document library name (optional, defaults to "Shared Documents")
    """
    try:
        ctx = get_sp_context_for_site(site_url)
        return get_folder_tree(parent_folder, ctx)
    except Exception as e:
        logger.error(f"Error getting folder tree: {str(e)}")
        return {"error": str(e)}

@mcp.tool(
    name="Get_Document_Content",
    description="Get content of a document in SharePoint with text extraction for Word/PDF/Excel. Requires site_url."
)
async def get_document_content_tool(site_url: str, folder_name: str, file_name: str, doc_library: Optional[str] = None):
    """
    Get content of a document in SharePoint with intelligent text extraction.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Relative folder path within doc library
        file_name: Name of the file to retrieve
        doc_library: Document library name (optional, defaults to "Shared Documents")
    """
    try:
        ctx = get_sp_context_for_site(site_url)
        return get_document_content(folder_name, file_name, ctx)
    except Exception as e:
        logger.error(f"Error getting document content: {str(e)}")
        return {"error": str(e)}

@mcp.tool(
    name="Create_Folder",
    description="Create a new folder in a SharePoint site. Requires site_url."
)
@_handle_sp_operation
async def create_folder(site_url: str, folder_name: str, parent_folder: Optional[str] = None, doc_library: Optional[str] = None):
    """
    Create a new folder in the specified directory or root if not specified.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Name of the new folder to create
        parent_folder: Parent folder path (optional, defaults to root)
        doc_library: Document library name (optional, defaults to "Shared Documents")
    """
    ctx = get_sp_context_for_site(site_url)
    parent_path = _get_path(parent_folder or "")
    logger.info(f"Creating folder '{folder_name}' in {parent_folder or 'root directory'} at {site_url}")

    # Check for existing folder
    if any(f["name"] == folder_name for f in list_folders(parent_folder, ctx)):
        return {"success": False, "message": f"Folder {folder_name} already exists"}

    # Create folder
    parent = ctx.web.get_folder_by_server_relative_url(parent_path)
    new_folder = parent.folders.add(folder_name)
    ctx.execute_query()

    return _file_success_response(new_folder, f"Folder {folder_name} created successfully")

@mcp.tool(
    name="Upload_Document",
    description="Upload a new file to a SharePoint directory. Requires site_url."
)
@_handle_sp_operation
async def upload_document(site_url: str, folder_name: str, file_name: str, content: str, is_base64: bool = False, doc_library: Optional[str] = None):
    """
    Upload a new file to a SharePoint directory.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Target folder path
        file_name: Name for the uploaded file
        content: File content (text or base64 encoded)
        is_base64: Whether content is base64 encoded (default: False)
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    logger.info(f"Uploading document {file_name} to folder {folder_name} at {site_url}")

    # Convert content and upload
    file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
    folder = ctx.web.get_folder_by_server_relative_url(_get_path(folder_name))
    uploaded_file = folder.upload_file(file_name, file_content)
    ctx.execute_query()

    return _file_success_response(uploaded_file, f"File {file_name} uploaded successfully")

@mcp.tool(
    name="Upload_Document_From_Path",
    description="Upload a file directly from local filesystem to SharePoint. Requires site_url."
)
@_handle_sp_operation
async def upload_document_from_path(site_url: str, folder_name: str, file_path: str, new_file_name: Optional[str] = None, doc_library: Optional[str] = None):
    """
    Upload a file directly from a path without needing to convert to base64 first.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Target folder path
        file_path: Local file path to upload
        new_file_name: Name for uploaded file (optional, uses original filename if not provided)
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    logger.info(f"Uploading document from path {file_path} to folder {folder_name} at {site_url}")

    try:
        with open(file_path, "rb") as file:
            file_content = file.read()

        if not new_file_name:
            new_file_name = os.path.basename(file_path)

        folder = ctx.web.get_folder_by_server_relative_url(_get_path(folder_name))
        uploaded_file = folder.upload_file(new_file_name, file_content)
        ctx.execute_query()

        return _file_success_response(uploaded_file, f"File {new_file_name} uploaded successfully")
    except Exception as e:
        logger.error(f"Error uploading file from path: {str(e)}")
        raise

@mcp.tool(
    name="Update_Document",
    description="Update an existing document in SharePoint. Requires site_url."
)
@_handle_sp_operation
async def update_document(site_url: str, folder_name: str, file_name: str, content: str, is_base64: bool = False, doc_library: Optional[str] = None):
    """
    Update an existing document in a SharePoint directory.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Folder path containing the file
        file_name: Name of the file to update
        content: New file content (text or base64 encoded)
        is_base64: Whether content is base64 encoded (default: False)
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    logger.info(f"Updating document {file_name} in folder {folder_name} at {site_url}")

    # Check if file exists
    file_path = _get_path(folder_name, file_name)
    file = ctx.web.get_file_by_server_relative_url(file_path)
    ctx.load(file, ["Exists", "Name", "ServerRelativeUrl"])
    ctx.execute_query()

    if not file.exists:
        return {"success": False, "message": f"File {file_name} does not exist in folder {folder_name}"}

    # Update file using upload method
    file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
    folder = ctx.web.get_folder_by_server_relative_url(_get_path(folder_name))
    updated_file = folder.upload_file(file_name, file_content)
    ctx.execute_query()

    return _file_success_response(updated_file, f"File {file_name} updated successfully")

@mcp.tool(
    name="Delete_Document",
    description="Delete a document from SharePoint. Requires site_url."
)
@_handle_sp_operation
async def delete_document(site_url: str, folder_name: str, file_name: str, doc_library: Optional[str] = None):
    """
    Delete a document from a SharePoint directory.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Folder path containing the file
        file_name: Name of the file to delete
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    logger.info(f"Deleting document {file_name} from folder {folder_name} at {site_url}")

    # Check if file exists and delete
    file = ctx.web.get_file_by_server_relative_url(_get_path(folder_name, file_name))
    ctx.load(file, ["Exists"])
    ctx.execute_query()

    if not file.exists:
        return {"success": False, "message": f"File {file_name} does not exist in folder {folder_name}"}

    file.delete_object()
    ctx.execute_query()
    return {"success": True, "message": f"File {file_name} deleted successfully"}

@mcp.tool(
    name="Delete_Folder",
    description="Delete an empty folder from SharePoint. Requires site_url."
)
@_handle_sp_operation
async def delete_folder(site_url: str, folder_path: str, doc_library: Optional[str] = None):
    """
    Delete an empty folder from SharePoint.

    Args:
        site_url: Full SharePoint site URL
        folder_path: Path of the folder to delete
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    logger.info(f"Deleting folder: {folder_path} at {site_url}")

    # Get folder and check if it exists and is empty
    full_path = _get_path(folder_path)
    folder = ctx.web.get_folder_by_server_relative_url(full_path)
    ctx.load(folder)
    ctx.load(folder.files)
    ctx.load(folder.folders)
    ctx.execute_query()

    if not hasattr(folder, 'exists') or not folder.exists:
        return {"success": False, "message": f"Folder '{folder_path}' does not exist"}

    if len(folder.files) > 0:
        return {"success": False, "message": f"Folder contains {len(folder.files)} files"}

    if len(folder.folders) > 0:
        return {"success": False, "message": f"Folder contains {len(folder.folders)} subfolders"}

    # Delete the empty folder
    folder.delete_object()
    ctx.execute_query()
    return {"success": True, "message": f"Folder '{folder_path}' deleted successfully"}

@mcp.tool(
    name="Download_Document",
    description="Download a document from SharePoint to local filesystem. Requires site_url."
)
@_handle_sp_operation
async def download_document_tool(site_url: str, folder_name: str, file_name: str, local_path: str, doc_library: Optional[str] = None):
    """
    Download a document from SharePoint to local filesystem with fallback support.

    Args:
        site_url: Full SharePoint site URL
        folder_name: Folder path containing the file
        file_name: Name of the file to download
        local_path: Local filesystem path for downloaded file
        doc_library: Document library name (optional)
    """
    ctx = get_sp_context_for_site(site_url)
    return download_document(folder_name, file_name, local_path, ctx)