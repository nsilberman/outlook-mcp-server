"""
Email operations for Outlook session management.

This module provides email-related operations such as moving emails between folders,
managing email policies, and retrieving email details.
"""

# Type imports
import os
import tempfile
from typing import Any, Dict, List, Optional, Tuple, Union

# Local application imports
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import email_cache, email_cache_order, get_email_from_cache
from ..validators import EmailNumberParam
from .exceptions import InvalidParameterError, OperationFailedError

logger = get_logger(__name__)


def _resolve_entry_id(email_identifier: Union[int, str]) -> Optional[str]:
    """Resolve an email identifier (position int or entry_id str) to an entry_id.

    Returns the entry_id string, or None if not found.
    """
    if isinstance(email_identifier, str):
        return email_identifier if email_identifier in email_cache else None
    # int path
    if not email_cache_order or email_identifier < 1 or email_identifier > len(email_cache_order):
        return None
    return email_cache_order[email_identifier - 1] or None


class EmailOperations:
    """Handles all email-related operations for Outlook."""
    
    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager

    def get_email_by_number(self, email_number: int) -> Dict[str, Any]:
        """
        Get email details by cache number.
        
        Args:
            email_number: The number of the email in the cache (1-based)
        
        Returns:
            Email dictionary with full details
        """
        try:
            # Validate input parameters
            if email_number < 1:
                raise ValueError(f"Invalid email number: {email_number}")
        except Exception as e:
            logger.error(f"Validation error in get_email_by_number: {e}")
            raise ValueError(f"Invalid parameters: {e}")

        try:
            if email_number not in email_cache:
                raise ValueError(f"Email #{email_number} not found in cache")
            
            # Get the entry_id from the cache
            entry_id = list(email_cache.keys())[email_number - 1] if email_number <= len(email_cache) else None
            if not entry_id:
                raise ValueError(f"Email #{email_number} not found in cache")
            
            email_data = email_cache[entry_id]
            logger.info(f"Retrieved email #{email_number}: {email_data.get('subject', 'No Subject')}")
            return email_data
            
        except Exception as e:
            error_msg = f"Error getting email by number: {e}"
            logger.error(error_msg)
            raise ValueError(error_msg)

    def move_email_to_folder(self, email_identifier: Union[int, str], target_folder_name: str) -> str:
        """
        Move an email to a different folder.

        Args:
            email_identifier: Position in cache (int, 1-based) or stable email ID (str)
            target_folder_name: Name of the target folder

        Returns:
            Success message or error message
        """
        if not target_folder_name or not target_folder_name.strip():
            return f"Error: Invalid folder name: {target_folder_name}"

        try:
            entry_id = _resolve_entry_id(email_identifier)
            if not entry_id:
                return f"Error: Email {email_identifier} not found in cache"
            
            with OutlookSessionManager() as session:
                # Get the email item
                item = session.namespace.GetItemFromID(entry_id)
                if not item:
                    return f"Error: Could not find email with entry ID {entry_id}"
                
                # Get target folder
                target_folder = session.get_folder(target_folder_name)
                if not target_folder:
                    return f"Error: Target folder '{target_folder_name}' not found"
                
                # Move the email
                item.Move(target_folder)
                
                # Remove from cache since it's been moved
                if entry_id in email_cache:
                    del email_cache[entry_id]
                    if entry_id in email_cache_order:
                        email_cache_order.remove(entry_id)
                
                logger.info(f"Moved email {email_identifier} to '{target_folder_name}'")
                return f"Email moved successfully to '{target_folder_name}'"

        except Exception as e:
            error_msg = f"Error moving email: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"

    def delete_email_by_number(self, email_identifier: Union[int, str]) -> str:
        """
        Delete an email by moving it to the Deleted Items folder.

        Args:
            email_identifier: Position in cache (int, 1-based) or stable email ID (str)

        Returns:
            Success message or error message
        """
        try:
            return self.move_email_to_folder(email_identifier, "Deleted Items")
            
        except Exception as e:
            error_msg = f"Error deleting email: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"


def get_email_by_number(email_number: int) -> Dict[str, Any]:
    """Get email details by cache number."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.get_email_by_number(email_number)


def move_email_to_folder(email_identifier: Union[int, str], target_folder_name: str) -> str:
    """Move an email to a different folder."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.move_email_to_folder(email_identifier, target_folder_name)


def delete_email_by_number(email_identifier: Union[int, str]) -> str:
    """Delete an email by moving it to the Deleted Items folder."""
    with OutlookSessionManager() as session_manager:
        email_ops = EmailOperations(session_manager)
        return email_ops.delete_email_by_number(email_identifier)


def get_email_categories(email_identifier: Union[int, str]) -> str:
    """Get the categories assigned to an email.

    Args:
        email_identifier: Position in cache (int, 1-based) or stable email ID (str)

    Returns:
        Comma-separated category names, or a message if none assigned
    """
    entry_id = _resolve_entry_id(email_identifier)
    if not entry_id:
        return f"Error: Email {email_identifier} not found in cache"

    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(entry_id)
            if not item:
                return f"Error: Could not find email with entry ID {entry_id}"

            categories = getattr(item, "Categories", "")
            if not categories or not categories.strip():
                return "No categories assigned"

            logger.info(f"Categories for email {email_identifier}: {categories}")
            return categories
        except Exception as e:
            error_msg = f"Error getting categories: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"


def set_email_categories(email_identifier: Union[int, str], categories: str) -> str:
    """Set or replace the categories on an email.

    Args:
        email_identifier: Position in cache (int, 1-based) or stable email ID (str)
        categories: Comma-separated category names (empty string to clear)

    Returns:
        Success or error message
    """
    entry_id = _resolve_entry_id(email_identifier)
    if not entry_id:
        return f"Error: Email {email_identifier} not found in cache"

    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(entry_id)
            if not item:
                return f"Error: Could not find email with entry ID {entry_id}"

            item.Categories = categories
            item.Save()

            if categories.strip():
                logger.info(f"Set categories for email {email_identifier}: {categories}")
                return f"Categories set to: {categories}"
            else:
                logger.info(f"Cleared categories for email {email_identifier}")
                return "Categories cleared"
        except Exception as e:
            error_msg = f"Error setting categories: {e}"
            logger.error(error_msg)
            return f"Error: {error_msg}"


# Image file extensions (count as 1 page)
_IMAGE_EXTENSIONS = frozenset(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp', '.svg', '.ico'))


def _count_pages(file_path: str) -> Optional[int]:
    """Count pages in a file. Returns None if format is unsupported.

    Supported: PDF (via pypdf), PPTX (via python-pptx), DOCX (via python-docx),
    and image files (always 1 page).
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in _IMAGE_EXTENSIONS:
        return 1

    if ext == '.pdf':
        try:
            from pypdf import PdfReader
            reader = PdfReader(file_path)
            return len(reader.pages)
        except ImportError:
            logger.warning("pypdf not installed — cannot count PDF pages")
            return None
        except Exception as e:
            logger.error(f"Error counting PDF pages: {e}")
            return None

    if ext == '.pptx':
        try:
            from pptx import Presentation
            prs = Presentation(file_path)
            return len(prs.slides)
        except ImportError:
            logger.warning("python-pptx not installed — cannot count PPTX slides")
            return None
        except Exception as e:
            logger.error(f"Error counting PPTX slides: {e}")
            return None

    if ext in ('.docx',):
        try:
            from docx import Document
            doc = Document(file_path)
            # DOCX has no native page count; estimate from section breaks + 1
            return len(doc.sections)
        except ImportError:
            logger.warning("python-docx not installed — cannot count DOCX pages")
            return None
        except Exception as e:
            logger.error(f"Error counting DOCX pages: {e}")
            return None

    return None


def save_attachment(email_identifier: Union[int, str], attachment_index: int, destination_dir: Optional[str] = None) -> Dict[str, Any]:
    """Save an attachment from a cached email to disk.

    Args:
        email_identifier: Position in cache (int, 1-based) or stable email ID (str)
        attachment_index: 1-based index of the attachment on the email
        destination_dir: Directory to save into (defaults to system temp dir)

    Returns:
        Dict with keys: success, file_path, file_name, size, error
    """
    entry_id = _resolve_entry_id(email_identifier)
    if not entry_id:
        return {"success": False, "error": f"Email {email_identifier} not found in cache"}

    save_dir = destination_dir or tempfile.gettempdir()

    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(entry_id)
            if not item:
                return {"success": False, "error": f"Could not find email with entry ID {entry_id}"}

            if not hasattr(item, 'Attachments') or not item.Attachments or item.Attachments.Count < attachment_index:
                return {"success": False, "error": f"Attachment #{attachment_index} not found (email has {getattr(item.Attachments, 'Count', 0)} attachments)"}

            attachment = item.Attachments.Item(attachment_index)
            file_name = getattr(attachment, 'FileName', None) or getattr(attachment, 'DisplayName', 'attachment')
            file_path = os.path.join(save_dir, file_name)
            attachment.SaveAsFile(file_path)

            size = os.path.getsize(file_path)
            logger.info(f"Saved attachment '{file_name}' from email {email_identifier} to {file_path}")
            return {"success": True, "file_path": file_path, "file_name": file_name, "size": size}
        except Exception as e:
            error_msg = f"Error saving attachment: {e}"
            logger.error(error_msg)
            return {"success": False, "error": error_msg}


def get_attachment_info(email_identifier: Union[int, str]) -> Dict[str, Any]:
    """Get detailed info about all attachments on a cached email, including page counts.

    Saves each attachment to a temp file, counts pages, then cleans up.

    Args:
        email_identifier: Position in cache (int, 1-based) or stable email ID (str)

    Returns:
        Dict with keys: success, attachments (list of dicts with name, size, pages), error
    """
    entry_id = _resolve_entry_id(email_identifier)
    if not entry_id:
        return {"success": False, "error": f"Email {email_identifier} not found in cache"}

    with OutlookSessionManager() as session:
        try:
            item = session.namespace.GetItemFromID(entry_id)
            if not item:
                return {"success": False, "error": f"Could not find email with entry ID {entry_id}"}

            if not hasattr(item, 'Attachments') or not item.Attachments or item.Attachments.Count == 0:
                return {"success": True, "attachments": []}

            results = []
            for i in range(1, item.Attachments.Count + 1):
                attachment = item.Attachments.Item(i)
                file_name = getattr(attachment, 'FileName', None) or getattr(attachment, 'DisplayName', 'attachment')
                size = getattr(attachment, 'Size', 0)

                # Try to count pages by saving to temp file
                pages = None
                ext = os.path.splitext(file_name)[1].lower()
                if ext in _IMAGE_EXTENSIONS:
                    pages = 1
                elif ext in ('.pdf', '.pptx', '.docx'):
                    tmp_path = None
                    try:
                        tmp_fd, tmp_path = tempfile.mkstemp(suffix=ext)
                        os.close(tmp_fd)
                        attachment.SaveAsFile(tmp_path)
                        pages = _count_pages(tmp_path)
                    except Exception as e:
                        logger.debug(f"Could not count pages for {file_name}: {e}")
                    finally:
                        if tmp_path and os.path.exists(tmp_path):
                            os.remove(tmp_path)

                info = {"index": i, "name": file_name, "size": size}
                if pages is not None:
                    info["pages"] = pages
                results.append(info)

            logger.info(f"Extracted info for {len(results)} attachments on email {email_identifier}")
            return {"success": True, "attachments": results}
        except Exception as e:
            error_msg = f"Error getting attachment info: {e}"
            logger.error(error_msg)
            return {"success": False, "error": error_msg}

