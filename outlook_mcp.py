#!/usr/bin/env python3
"""
Complete Microsoft Graph MCP Server for Claude Desktop
Includes all Graph API functionality in one file
"""

import asyncio
import json
import os
import sys
import requests
import base64
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
from urllib.parse import urlencode

# MCP imports
from mcp.server import Server
from mcp.server.models import InitializationOptions
import mcp.server.stdio
import mcp.types as types

class GraphClient:
    """Combined Graph API client for email and calendar"""
    
    def __init__(self, client_id, client_secret, tenant_id="common"):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.base_url = "https://graph.microsoft.com/v1.0"
    
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
    
    # EMAIL METHODS
    def get_messages(self, folder_id="inbox", top=50, search=None, filter_str=None):
        """Get email messages"""
        endpoint = f"/me/mailFolders/{folder_id}/messages"
        
        params = {
            "$top": top,
            "$orderby": "receivedDateTime desc"
        }
        
        if search:
            params["$search"] = f'"{search}"'
        
        if filter_str:
            params["$filter"] = filter_str
        
        return self._make_request("GET", endpoint, params=params)
    
    # def send_email(self, to_recipients, subject, body, cc_recipients=None, body_type="HTML"):
    #    """Send an email (COMMENTED OUT - READ ONLY MODE)"""
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
    #    data = {
    #        "message": message,
    #        "saveToSentItems": True
    #    }
    #    
    #    return self._make_request("POST", "/me/sendMail", data)
    
    def search_messages(self, query, top=25):
        """Search messages across all folders"""
        return self.get_messages(folder_id="inbox", search=query, top=top)
    
    # CALENDAR METHODS
    def get_events(self, start_date=None, end_date=None, top=50):
        """Get calendar events"""
        endpoint = "/me/events"
        
        params = {"$top": top, "$orderby": "start/dateTime"}
        
        if start_date and end_date:
            params["$filter"] = f"start/dateTime ge '{start_date}' and end/dateTime le '{end_date}'"
        
        return self._make_request("GET", endpoint, params=params)
    
    def create_event(self, subject, start_time, end_time, body=None, location=None, 
                    attendees=None, timezone_name="UTC"):
        """Create a new calendar event"""
        event_data = {
            "subject": subject,
            "start": {
                "dateTime": start_time.isoformat(),
                "timeZone": timezone_name
            },
            "end": {
                "dateTime": end_time.isoformat(),
                "timeZone": timezone_name
            }
        }
        
        if body:
            event_data["body"] = {
                "contentType": "text",
                "content": body
            }
        
        if location:
            event_data["location"] = {
                "displayName": location
            }
        
        if attendees:
            event_data["attendees"] = [
                {
                    "emailAddress": {
                        "address": email,
                        "name": email.split("@")[0]
                    },
                    "type": "required"
                }
                for email in attendees
            ]
        
        return self._make_request("POST", "/me/events", event_data)
    
    def get_user_info(self):
        """Get current user information"""
        return self._make_request("GET", "/me")

# Initialize MCP server
server = Server("outlook-connector")

# Global client instance
graph_client = None

def initialize_graph_client():
    """Initialize the Graph client with stored credentials"""
    global graph_client
    
    # Get credentials from environment variables
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    
    if not client_id or not client_secret:
        print("❌ Missing AZURE_CLIENT_ID or AZURE_CLIENT_SECRET environment variables", file=sys.stderr)
        return False
    
    graph_client = GraphClient(client_id, client_secret)
    
    # Load saved access token if available
    token_file = os.path.expanduser("~/.outlook_token.json")
    if os.path.exists(token_file):
        try:
            with open(token_file, 'r') as f:
                token_data = json.load(f)
                graph_client.access_token = token_data.get("access_token")
                print("✅ Loaded saved access token", file=sys.stderr)
                return True
        except Exception as e:
            print(f"⚠️  Could not load saved token: {e}", file=sys.stderr)
    
    print("⚠️  No valid access token found. Please run authentication first.", file=sys.stderr)
    return False

@server.list_tools()
async def handle_list_tools() -> List[types.Tool]:
    """List available tools"""
    return [
        types.Tool(
            name="get_emails",
            description="Get recent emails from inbox or search emails",
            inputSchema={
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": "Number of emails to retrieve (default: 10, max: 50)",
                        "default": 10,
                        "maximum": 50
                    },
                    "search": {
                        "type": "string",
                        "description": "Search query to filter emails (optional)"
                    }
                }
            }
        ),
        # types.Tool(
        #     name="send_email",
        #     description="Send an email message (DISABLED - READ ONLY MODE)",
        #     inputSchema={
        #         "type": "object",
        #         "properties": {
        #             "to": {
        #                 "type": "array",
        #                 "items": {"type": "string"},
        #                 "description": "Recipient email addresses"
        #             },
        #             "subject": {
        #                 "type": "string",
        #                 "description": "Email subject"
        #             },
        #             "body": {
        #                 "type": "string",
        #                 "description": "Email body content (supports HTML)"
        #             },
        #             "cc": {
        #                 "type": "array",
        #                 "items": {"type": "string"},
        #                 "description": "CC recipients (optional)"
        #             }
        #         },
        #         "required": ["to", "subject", "body"]
        #     }
        # ),
        types.Tool(
            name="get_calendar_events",
            description="Get upcoming calendar events",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {
                        "type": "integer",
                        "description": "Number of days to look ahead (default: 7)",
                        "default": 7,
                        "maximum": 30
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum number of events (default: 20)",
                        "default": 20,
                        "maximum": 50
                    }
                }
            }
        ),
        types.Tool(
            name="create_calendar_event",
            description="Create a new calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {
                        "type": "string",
                        "description": "Event title"
                    },
                    "start_time": {
                        "type": "string",
                        "description": "Start time in ISO format (e.g., 2025-08-14T14:00:00) or natural language"
                    },
                    "end_time": {
                        "type": "string",
                        "description": "End time in ISO format or natural language"
                    },
                    "location": {
                        "type": "string",
                        "description": "Event location (optional)"
                    },
                    "description": {
                        "type": "string",
                        "description": "Event description (optional)"
                    },
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Attendee email addresses (optional)"
                    }
                },
                "required": ["subject", "start_time", "end_time"]
            }
        ),
        types.Tool(
            name="get_user_info",
            description="Get current user's profile information",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[types.TextContent]:
    """Handle tool calls"""
    
    if not graph_client or not graph_client.access_token:
        return [types.TextContent(
            type="text",
            text="❌ Not authenticated with Microsoft Graph. Please run the authentication script first and ensure your access token is saved."
        )]
    
    try:
        if name == "get_emails":
            count = min(arguments.get("count", 10), 50)
            search = arguments.get("search")
            
            if search:
                result = graph_client.search_messages(search, top=count)
            else:
                result = graph_client.get_messages(top=count)
            
            emails = result.get("value", [])
            
            if not emails:
                search_text = f" matching '{search}'" if search else ""
                return [types.TextContent(
                    type="text",
                    text=f"📧 No emails found{search_text}."
                )]
            
            email_summaries = []
            for email in emails:
                from_addr = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                subject = email.get("subject", "No subject")
                received = email.get("receivedDateTime", "Unknown")
                is_read = email.get("isRead", False)
                preview = email.get("bodyPreview", "")
                
                # Format date
                try:
                    dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                    formatted_date = dt.strftime("%Y-%m-%d %H:%M")
                except:
                    formatted_date = received
                
                email_summaries.append(
                    f"📧 **{subject}**\n"
                    f"   From: {from_addr}\n"
                    f"   Date: {formatted_date}\n"
                    f"   Status: {'✅ Read' if is_read else '🔴 Unread'}\n"
                    f"   Preview: {preview[:150]}{'...' if len(preview) > 150 else ''}\n"
                )
            
            search_text = f" matching '{search}'" if search else ""
            return [types.TextContent(
                type="text",
                text=f"📧 Found {len(emails)} emails{search_text}:\n\n" + "\n".join(email_summaries)
            )]
        
        # elif name == "send_email":
        #     # EMAIL SENDING DISABLED - READ ONLY MODE
        #     return [types.TextContent(
        #         type="text",
        #         text="❌ Email sending is disabled. This MCP is in read-only mode."
        #     )]
        
        elif name == "get_calendar_events":
            days = min(arguments.get("days", 7), 30)
            count = min(arguments.get("count", 20), 50)
            
            # Calculate date range
            start_date = datetime.now().isoformat() + "Z"
            end_date = (datetime.now() + timedelta(days=days)).isoformat() + "Z"
            
            result = graph_client.get_events(
                start_date=start_date,
                end_date=end_date,
                top=count
            )
            
            events = result.get("value", [])
            
            if not events:
                return [types.TextContent(
                    type="text",
                    text=f"📅 No events found in the next {days} days."
                )]
            
            event_summaries = []
            for event in events:
                subject = event.get("subject", "No title")
                start_dt = event.get("start", {}).get("dateTime", "")
                end_dt = event.get("end", {}).get("dateTime", "")
                location = event.get("location", {}).get("displayName", "No location")
                attendees = event.get("attendees", [])
                
                # Format dates
                try:
                    start = datetime.fromisoformat(start_dt.replace('Z', '+00:00'))
                    end = datetime.fromisoformat(end_dt.replace('Z', '+00:00'))
                    
                    if start.date() == end.date():
                        time_str = f"{start.strftime('%Y-%m-%d %H:%M')} - {end.strftime('%H:%M')}"
                    else:
                        time_str = f"{start.strftime('%Y-%m-%d %H:%M')} - {end.strftime('%Y-%m-%d %H:%M')}"
                except:
                    time_str = f"{start_dt} - {end_dt}"
                
                attendee_list = [att.get("emailAddress", {}).get("address", "") for att in attendees]
                attendee_text = f"\n   Attendees: {', '.join(attendee_list)}" if attendee_list else ""
                
                event_summaries.append(
                    f"📅 **{subject}**\n"
                    f"   When: {time_str}\n"
                    f"   Where: {location}{attendee_text}\n"
                )
            
            return [types.TextContent(
                type="text",
                text=f"📅 Found {len(events)} events in the next {days} days:\n\n" + "\n".join(event_summaries)
            )]
        
        elif name == "create_calendar_event":
            subject = arguments["subject"]
            start_time_str = arguments["start_time"]
            end_time_str = arguments["end_time"]
            location = arguments.get("location")
            description = arguments.get("description")
            attendees = arguments.get("attendees")
            
            # Parse datetime strings
            try:
                # Handle various datetime formats
                for fmt in ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M"]:
                    try:
                        start_time = datetime.strptime(start_time_str.replace('Z', ''), fmt)
                        break
                    except ValueError:
                        continue
                else:
                    raise ValueError(f"Could not parse start time: {start_time_str}")
                
                for fmt in ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M"]:
                    try:
                        end_time = datetime.strptime(end_time_str.replace('Z', ''), fmt)
                        break
                    except ValueError:
                        continue
                else:
                    raise ValueError(f"Could not parse end time: {end_time_str}")
                
            except ValueError as e:
                return [types.TextContent(
                    type="text",
                    text=f"❌ Error parsing datetime: {e}\n"
                         f"Please use format like: 2025-08-14T14:00:00"
                )]
            
            result = graph_client.create_event(
                subject=subject,
                start_time=start_time,
                end_time=end_time,
                location=location,
                body=description,
                attendees=attendees
            )
            
            attendee_text = f"\n   Attendees: {', '.join(attendees)}" if attendees else ""
            location_text = f"\n   Location: {location}" if location else ""
            
            return [types.TextContent(
                type="text",
                text=f"✅ Calendar event created successfully!\n"
                     f"   Title: {subject}\n"
                     f"   When: {start_time.strftime('%Y-%m-%d %H:%M')} - {end_time.strftime('%H:%M')}{location_text}{attendee_text}"
            )]
        
        elif name == "get_user_info":
            user = graph_client.get_user_info()
            
            return [types.TextContent(
                type="text",
                text=f"👤 **User Information**\n"
                     f"   Name: {user.get('displayName', 'N/A')}\n"
                     f"   Email: {user.get('mail', user.get('userPrincipalName', 'N/A'))}\n"
                     f"   Job Title: {user.get('jobTitle', 'N/A')}\n"
                     f"   Office: {user.get('officeLocation', 'N/A')}\n"
                     f"   Phone: {user.get('businessPhones', ['N/A'])[0] if user.get('businessPhones') else 'N/A'}"
            )]
        
        else:
            return [types.TextContent(
                type="text",
                text=f"❌ Unknown tool: {name}"
            )]
    
    except Exception as e:
        return [types.TextContent(
            type="text",
            text=f"❌ Error executing {name}: {str(e)}"
        )]

async def main():
    """Main entry point"""
    try:
        # Initialize the Graph client
        if not initialize_graph_client():
            print("❌ Failed to initialize Graph client", file=sys.stderr)
            sys.exit(1)
        
        # Run the MCP server
        async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
            await server.run(
                read_stream,
                write_stream,
                InitializationOptions(
                    server_name="outlook-connector",
                    server_version="1.0.0",
                    capabilities=server.get_capabilities(
                        notification_options=None,
                        experimental_capabilities=None
                    )
                )
            )
    except Exception as e:
        print(f"❌ Server error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(main())
