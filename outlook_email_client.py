import requests
import json
import base64
from datetime import datetime
from typing import List, Dict, Optional, Union
import mimetypes
import os
from dotenv import load_dotenv

class GraphEmailClient:
    def __init__(self, client_id, client_secret, tenant_id=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id or "common"
        self.access_token = None
        self.base_url = "https://graph.microsoft.com/v1.0"
    
    def get_auth_url(self, redirect_uri, scopes=None):
        """Generate authorization URL with email scopes"""
        if scopes is None:
            scopes = [
                "Mail.Read", 
                "Mail.ReadWrite", 
                "Mail.Send", 
                "User.Read",
                "Calendars.ReadWrite"  # Keep calendar access too
            ]
        
        from urllib.parse import urlencode
        params = {
            "client_id": self.client_id,
            "response_type": "code",
            "redirect_uri": redirect_uri,
            "scope": " ".join(scopes),
            "response_mode": "query"
        }
        
        auth_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize"
        return f"{auth_url}?{urlencode(params)}"
    
    def get_access_token(self, authorization_code, redirect_uri):
        """Exchange authorization code for access token"""
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "code": authorization_code,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code"
        }
        
        response = requests.post(token_url, data=data)
        
        if response.status_code == 200:
            token_data = response.json()
            self.access_token = token_data["access_token"]
            return token_data
        else:
            raise Exception(f"Token request failed: {response.text}")
    
    def _make_request(self, method, endpoint, data=None, params=None):
        """Make authenticated request to Graph API"""
        if not self.access_token:
            raise Exception("No access token available. Please authenticate first.")
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"{self.base_url}{endpoint}"
        
        kwargs = {"headers": headers}
        if data:
            kwargs["json"] = data
        if params:
            kwargs["params"] = params
        
        response = getattr(requests, method.lower())(url, **kwargs)
        
        if response.status_code in [200, 201, 202, 204]:
            return response.json() if response.content else None
        else:
            raise Exception(f"API request failed: {response.status_code} - {response.text}")
    
    # EMAIL READING METHODS
    
    def get_messages(self, folder_id="inbox", top=50, skip=0, search=None, 
                    filter_str=None, select_fields=None, order_by="receivedDateTime desc"):
        """Get email messages from a folder"""
        endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": order_by
        }
        
        if search:
            params["$search"] = f'"{search}"'
        
        if filter_str:
            params["$filter"] = filter_str
        
        if select_fields:
            params["$select"] = ",".join(select_fields)
        
        return self._make_request("GET", endpoint, params=params)
    
    def get_message(self, message_id, select_fields=None):
        """Get a specific email message"""
        endpoint = f"/me/messages/{message_id}"
        
        params = {}
        if select_fields:
            params["$select"] = ",".join(select_fields)
        
        return self._make_request("GET", endpoint, params=params)
    
    def get_message_attachments(self, message_id):
        """Get attachments for a message"""
        endpoint = f"/me/messages/{message_id}/attachments"
        return self._make_request("GET", endpoint)
    
    def download_attachment(self, message_id, attachment_id, save_path=None):
        """Download an email attachment"""
        endpoint = f"/me/messages/{message_id}/attachments/{attachment_id}"
        attachment = self._make_request("GET", endpoint)
        
        if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
            content = base64.b64decode(attachment["contentBytes"])
            
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(content)
                return save_path
            else:
                return content
        else:
            raise Exception("Unsupported attachment type")
    
    def search_messages(self, query, top=25):
        """Search messages across all folders"""
        return self.get_messages(folder_id="inbox", search=query, top=top)
    
    # EMAIL FOLDER METHODS
    
    def get_mail_folders(self):
        """Get all mail folders"""
        return self._make_request("GET", "/me/mailFolders")
    
    def create_folder(self, display_name, parent_folder_id=None):
        """Create a new mail folder"""
        endpoint = f"/me/mailFolders/{parent_folder_id}/childFolders" if parent_folder_id else "/me/mailFolders"
        
        data = {"displayName": display_name}
        return self._make_request("POST", endpoint, data)
    
    # EMAIL SENDING METHODS (COMMENTED OUT - READ ONLY MODE)
    
    # def send_email(self, to_recipients, subject, body, cc_recipients=None, 
    #              bcc_recipients=None, attachments=None, body_type="HTML", 
    #              importance="normal", save_to_sent_items=True):
    #    """Send an email message"""
    #    
    #    # Build recipient lists
    #    def build_recipients(emails):
    #        if isinstance(emails, str):
    #            emails = [emails]
    #        return [{"emailAddress": {"address": email}} for email in emails]
    #    
    #    message = {
    #        "subject": subject,
    #        "body": {
    #            "contentType": body_type,
    #            "content": body
    #        },
    #        "toRecipients": build_recipients(to_recipients),
    #        "importance": importance
    #    }
    #    
    #    if cc_recipients:
    #        message["ccRecipients"] = build_recipients(cc_recipients)
    #    
    #    if bcc_recipients:
    #        message["bccRecipients"] = build_recipients(bcc_recipients)
    #    
    #    if attachments:
    #        message["attachments"] = self._prepare_attachments(attachments)
    #    
    #    data = {
    #        "message": message,
    #        "saveToSentItems": save_to_sent_items
    #    }
    #    
    #    return self._make_request("POST", "/me/sendMail", data)
    
    # def create_draft(self, to_recipients, subject, body, cc_recipients=None, 
    #                bcc_recipients=None, attachments=None, body_type="HTML"):
    #    """Create a draft email"""
    #    
    #    def build_recipients(emails):
    #        if isinstance(emails, str):
    #            emails = [emails]
    #        return [{"emailAddress": {"address": email}} for email in emails]
    #    
    #    message = {
    #        "subject": subject,
    #        "body": {
    #            "contentType": body_type,
    #            "content": body
    #        },
    #        "toRecipients": build_recipients(to_recipients)
    #    }
    #    
    #    if cc_recipients:
    #        message["ccRecipients"] = build_recipients(cc_recipients)
    #    
    #    if bcc_recipients:
    #        message["bccRecipients"] = build_recipients(bcc_recipients)
    #    
    #    if attachments:
    #        message["attachments"] = self._prepare_attachments(attachments)
    #    
    #    return self._make_request("POST", "/me/messages", message)
    
    # def send_draft(self, message_id):
    #    """Send a draft message"""
    #    return self._make_request("POST", f"/me/messages/{message_id}/send")
    
    # def _prepare_attachments(self, attachments):
    #    """Prepare attachments for email sending"""
    #    prepared = []
    #    
    #    for attachment in attachments:
    #        if isinstance(attachment, str):
    #            # File path
    #            with open(attachment, "rb") as f:
    #                content = base64.b64encode(f.read()).decode()
    #            
    #            filename = os.path.basename(attachment)
    #            content_type = mimetypes.guess_type(attachment)[0] or "application/octet-stream"
    #            
    #            prepared.append({
    #                "@odata.type": "#microsoft.graph.fileAttachment",
    #                "name": filename,
    #                "contentBytes": content,
    #                "contentType": content_type
    #            })
    #        
    #        elif isinstance(attachment, dict):
    #            # Dictionary with file info
    #            prepared.append({
    #                "@odata.type": "#microsoft.graph.fileAttachment",
    #                "name": attachment["name"],
    #                "contentBytes": attachment["content"],
    #                "contentType": attachment.get("content_type", "application/octet-stream")
    #            })
    #    
    #    return prepared
    
    # EMAIL MANAGEMENT METHODS
    
    def mark_as_read(self, message_id):
        """Mark a message as read"""
        data = {"isRead": True}
        return self._make_request("PATCH", f"/me/messages/{message_id}", data)
    
    def mark_as_unread(self, message_id):
        """Mark a message as unread"""
        data = {"isRead": False}
        return self._make_request("PATCH", f"/me/messages/{message_id}", data)
    
    def delete_message(self, message_id):
        """Delete a message (move to deleted items)"""
        return self._make_request("DELETE", f"/me/messages/{message_id}")
    
    def move_message(self, message_id, destination_folder_id):
        """Move a message to a different folder"""
        data = {"destinationId": destination_folder_id}
        return self._make_request("POST", f"/me/messages/{message_id}/move", data)
    
    def copy_message(self, message_id, destination_folder_id):
        """Copy a message to a different folder"""
        data = {"destinationId": destination_folder_id}
        return self._make_request("POST", f"/me/messages/{message_id}/copy", data)
    
    def flag_message(self, message_id, flag_status="flagged"):
        """Flag a message"""
        data = {
            "flag": {
                "flagStatus": flag_status
            }
        }
        return self._make_request("PATCH", f"/me/messages/{message_id}", data)
    
    # RULE MANAGEMENT
    
    def get_message_rules(self):
        """Get all message rules"""
        return self._make_request("GET", "/me/mailFolders/inbox/messageRules")
    
    def create_message_rule(self, display_name, conditions, actions, enabled=True):
        """Create a new message rule"""
        rule = {
            "displayName": display_name,
            "sequence": 1,
            "isEnabled": enabled,
            "conditions": conditions,
            "actions": actions
        }
        
        return self._make_request("POST", "/me/mailFolders/inbox/messageRules", rule)
    
    # UTILITY METHODS
    
    def get_unread_count(self, folder_id="inbox"):
        """Get count of unread messages in a folder"""
        folder = self._make_request("GET", f"/me/mailFolders/{folder_id}")
        return folder.get("unreadItemCount", 0)
    
    def get_recent_emails(self, hours=24, folder_id="inbox"):
        """Get emails from the last N hours"""
        from datetime import datetime, timedelta
        
        cutoff_time = (datetime.now() - timedelta(hours=hours)).isoformat() + "Z"
        filter_str = f"receivedDateTime ge {cutoff_time}"
        
        return self.get_messages(folder_id=folder_id, filter_str=filter_str)
    
    def bulk_mark_read(self, message_ids):
        """Mark multiple messages as read"""
        results = []
        for msg_id in message_ids:
            try:
                result = self.mark_as_read(msg_id)
                results.append({"id": msg_id, "success": True, "result": result})
            except Exception as e:
                results.append({"id": msg_id, "success": False, "error": str(e)})
        return results


# Usage example
def email_example():
    """Example usage of the email connector"""
    # Load environment variables
    load_dotenv()
    
    CLIENT_ID = os.getenv("CLIENT_ID")
    CLIENT_SECRET = os.getenv("CLIENT_SECRET")
    REDIRECT_URI = "http://localhost:8000/callback"
    
    if not CLIENT_ID or not CLIENT_SECRET:
        print("Error: CLIENT_ID and CLIENT_SECRET must be set in .env file")
        return
    
    client = GraphEmailClient(CLIENT_ID, CLIENT_SECRET)
    
    # Authentication (same as calendar example)
    auth_url = client.get_auth_url(REDIRECT_URI)
    print(f"Visit: {auth_url}")
    auth_code = input("Enter authorization code: ")
    
    client.get_access_token(auth_code, REDIRECT_URI)
    
    # Get recent emails
    recent_emails = client.get_messages(top=10)
    print(f"Found {len(recent_emails['value'])} recent emails:")
    
    for email in recent_emails['value']:
        print(f"- From: {email['from']['emailAddress']['address']}")
        print(f"  Subject: {email['subject']}")
        print(f"  Received: {email['receivedDateTime']}")
        print(f"  Read: {email['isRead']}")
        print()
    
    # Send a test email (COMMENTED OUT - READ ONLY MODE)
    # client.send_email(
    #     to_recipients=["recipient@example.com"],
    #     subject="Test Email from Graph API",
    #     body="<h1>Hello!</h1><p>This email was sent using Microsoft Graph API.</p>",
    #     body_type="HTML"
    # )
    # print("Test email sent!")
    
    # Search for emails
    search_results = client.search_messages("meeting")
    print(f"Found {len(search_results['value'])} emails containing 'meeting'")
    
    # Get unread count
    unread_count = client.get_unread_count()
    print(f"You have {unread_count} unread emails")


if __name__ == "__main__":
    email_example()
