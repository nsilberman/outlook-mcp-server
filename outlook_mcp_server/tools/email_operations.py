"""Email operations tools for Outlook MCP Server."""

from typing import Dict, Any, Union, List, Optional
from ..backend.email_composition import reply_to_email_by_number, compose_email, create_draft
from ..backend.outlook_session import OutlookSessionManager
from ..backend.validation import ValidationError


def reply_to_email_by_number_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Union[str, List[str], None] = None,
    cc_recipients: Union[str, List[str], None] = None,
) -> Dict[str, Any]:
    """Send a reply to an email immediately. Use create_reply_draft_tool instead to save as draft first.

    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]

    Behavior:
        - When both to_recipients and cc_recipients are None:
          * Uses ReplyAll() to maintain original recipients
        - When either parameter is provided:
          * Uses Reply() with specified recipients
          * Any None parameters will result in empty recipient fields
        - Single email strings and lists of email strings are both accepted

    Returns:
        dict: Response containing confirmation message
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")

    try:
        result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients, save_as_draft=False)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error replying to email: {str(e)}"}


def create_reply_draft_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Union[str, List[str], None] = None,
    cc_recipients: Union[str, List[str], None] = None,
    html: bool = False,
) -> Dict[str, Any]:
    """Prepare a reply to an email and save it as a draft (does NOT send). Use reply_to_email_by_number_tool to send immediately.

    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        html: If True, reply_text is treated as HTML (default: False)

    Returns:
        dict: Response containing confirmation message
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")

    try:
        result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients, save_as_draft=True, html=html)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error creating reply draft: {str(e)}"}


def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> Dict[str, Any]:
    """Compose and send a new email

    Args:
        recipient_email: Email address(es) of the recipient(s) - can be single email or semicolon-separated list
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address(es) - can be single email or semicolon-separated list

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")
    
    try:
        # Parse semicolon-separated email addresses into lists
        to_recipients = [email.strip() for email in recipient_email.split(';') if email.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [email.strip() for email in cc_email.split(';') if email.strip()]
        
        result = compose_email(to_recipients, subject, body, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error composing email: {str(e)}"}


def move_email_tool(email_number: int, target_folder_name: str) -> Dict[str, Any]:
    """Move an email to the specified folder.

    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        After moving, the cache will be cleared to reflect the new email positions.
        
        IMPORTANT: Target folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValidationError("Target folder name must be a non-empty string")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import move_email_to_folder
        result = move_email_to_folder(email_number, target_folder_name)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


def delete_email_by_number_tool(email_number: int) -> Dict[str, Any]:
    """Move an email to the Deleted Items folder.

    Args:
        email_number: The number of the email in the cache to delete (1-based)

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved to Deleted Items successfully"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        This tool moves the email to the Deleted Items folder instead of permanently deleting it.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import delete_email_by_number
        result = delete_email_by_number(email_number)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting email: {str(e)}"}


def create_draft_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None, html: bool = False, attachments: Optional[str] = None) -> Dict[str, Any]:
    """Create a draft email without sending it

    Args:
        recipient_email: Email address(es) of the recipient(s) - can be single email or semicolon-separated list
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address(es) - can be single email or semicolon-separated list
        html: If True, body is treated as HTML (default: False)
        attachments: Optional file path(s) to attach - semicolon-separated for multiple files (e.g. "C:\\report.pdf;C:\\data.xlsx")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")

    try:
        # Parse semicolon-separated email addresses into lists
        to_recipients = [email.strip() for email in recipient_email.split(';') if email.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [email.strip() for email in cc_email.split(';') if email.strip()]

        # Parse attachments: accept semicolon-separated string or JSON array string
        attachment_list = None
        if attachments:
            if isinstance(attachments, list):
                attachment_list = attachments
            elif isinstance(attachments, str):
                # Try JSON array first (e.g. '["path1", "path2"]')
                import json
                try:
                    parsed = json.loads(attachments)
                    if isinstance(parsed, list):
                        attachment_list = [str(p).strip() for p in parsed if str(p).strip()]
                    else:
                        attachment_list = [attachments.strip()]
                except (json.JSONDecodeError, ValueError):
                    # Fall back to semicolon-separated
                    attachment_list = [p.strip() for p in attachments.split(';') if p.strip()]

        result = create_draft(to_recipients, subject, body, cc_recipients, html, attachment_list)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error creating draft: {str(e)}"}


def get_email_categories_tool(email_number: int) -> Dict[str, Any]:
    """Get the categories assigned to an email

    Args:
        email_number: The number of the email in the cache (1-based)

    Returns:
        dict: Response containing the comma-separated category names
        {
            "type": "text",
            "text": "Category1, Category2"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    try:
        from ..backend.outlook_session.email_operations import get_email_categories
        result = get_email_categories(email_number)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error getting categories: {str(e)}"}


def set_email_categories_tool(email_number: int, categories: str) -> Dict[str, Any]:
    """Set or replace the categories on an email

    Args:
        email_number: The number of the email in the cache (1-based)
        categories: Comma-separated category names to assign (e.g. "Important, Follow-up"). Use empty string to clear all categories.

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Categories set to: Important, Follow-up"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        This replaces all existing categories. To add a category, first read existing
        categories with get_email_categories_tool, then include them in the new value.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not isinstance(categories, str):
        raise ValidationError("Categories must be a string")

    try:
        from ..backend.outlook_session.email_operations import set_email_categories
        result = set_email_categories(email_number, categories)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error setting categories: {str(e)}"}


def get_attachment_info_tool(email_number: int) -> Dict[str, Any]:
    """Get detailed information about all attachments on an email, including page counts

    Returns each attachment's name, size, index, and page count when available.
    Page counts are supported for PDF, PPTX, DOCX, and image files (images = 1 page).
    Requires optional dependencies: pypdf, python-pptx, python-docx.

    Args:
        email_number: The number of the email in the cache (1-based)

    Returns:
        dict: Response containing attachment details as JSON
        {
            "type": "json",
            "data": {
                "attachments": [
                    {"index": 1, "name": "report.pdf", "size": 102400, "pages": 12},
                    {"index": 2, "name": "photo.png", "size": 54321, "pages": 1}
                ]
            }
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    try:
        from ..backend.outlook_session.email_operations import get_attachment_info
        result = get_attachment_info(email_number)
        if result["success"]:
            return {"type": "json", "data": {"attachments": result["attachments"]}}
        return {"type": "text", "text": f"Error: {result['error']}"}
    except Exception as e:
        return {"type": "text", "text": f"Error getting attachment info: {str(e)}"}


def save_attachment_tool(email_number: int, attachment_index: int, destination_dir: Optional[str] = None) -> Dict[str, Any]:
    """Save an attachment from an email to disk

    Args:
        email_number: The number of the email in the cache (1-based)
        attachment_index: 1-based index of the attachment (use get_attachment_info_tool to see indices)
        destination_dir: Optional directory path to save the file (defaults to system temp directory)

    Returns:
        dict: Response containing the saved file path
        {
            "type": "text",
            "text": "Saved 'report.pdf' (102400 bytes) to /tmp/report.pdf"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not isinstance(attachment_index, int) or attachment_index < 1:
        raise ValidationError("Attachment index must be a positive integer")

    try:
        from ..backend.outlook_session.email_operations import save_attachment
        result = save_attachment(email_number, attachment_index, destination_dir)
        if result["success"]:
            return {"type": "text", "text": f"Saved '{result['file_name']}' ({result['size']} bytes) to {result['file_path']}"}
        return {"type": "text", "text": f"Error: {result['error']}"}
    except Exception as e:
        return {"type": "text", "text": f"Error saving attachment: {str(e)}"}