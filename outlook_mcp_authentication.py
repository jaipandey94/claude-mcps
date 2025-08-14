#!/usr/bin/env python3
"""
One-time authentication script for MCP Outlook server
Run this once to authenticate and save your access token
"""

import requests
import webbrowser
import json
import os
from urllib.parse import urlencode
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# CONFIGURATION - Load from environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = "http://localhost:8000/callback"
TENANT_ID = "common"  # Use "consumers" if you only want personal accounts

def authenticate():
    """Perform one-time authentication and save token"""
    
    print("🔐 Microsoft Graph Authentication for MCP Server")
    print("=" * 50)
    
    # Validate configuration
    if not CLIENT_ID or not CLIENT_SECRET:
        print("❌ ERROR: CLIENT_ID and CLIENT_SECRET must be set in .env file")
        print("   Please check your .env file contains:")
        print("   CLIENT_ID=your-actual-client-id")
        print("   CLIENT_SECRET=your-actual-client-secret")
        return False
    
    # Step 1: Generate authorization URL
    scopes = [
        "User.Read",
        "Calendars.ReadWrite", 
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send"
    ]
    
    auth_params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "scope": " ".join(scopes),
        "response_mode": "query"
    }
    
    auth_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
    full_auth_url = f"{auth_url}?{urlencode(auth_params)}"
    
    print("🔗 Opening browser for authentication...")
    print(f"URL: {full_auth_url}")
    print()
    
    try:
        webbrowser.open(full_auth_url)
    except:
        print("Could not open browser automatically. Please copy the URL above.")
    
    print("📝 Instructions:")
    print("1. Sign in with your Microsoft account")
    print("2. Accept the permissions")
    print("3. Copy the 'code' parameter from the URL after redirect")
    print()
    
    # Get authorization code
    auth_code = input("📥 Paste the authorization code here: ").strip()
    
    if not auth_code:
        print("❌ No authorization code provided")
        return False
    
    # Step 2: Exchange code for token
    print("🔄 Exchanging code for access token...")
    
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": auth_code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }
    
    try:
        response = requests.post(token_url, data=token_data)
        
        if response.status_code == 200:
            token_info = response.json()
            
            # Save token to file
            token_file = os.path.expanduser("~/.outlook_token.json")
            with open(token_file, 'w') as f:
                json.dump(token_info, f, indent=2)
            
            print("✅ Authentication successful!")
            print(f"✅ Token saved to: {token_file}")
            print(f"✅ Token expires in: {token_info.get('expires_in', 'unknown')} seconds")
            
            # Test the token
            print("\n🧪 Testing API access...")
            test_token(token_info["access_token"])
            
            # Set up environment variables
            print("\n📝 Next Steps:")
            print("1. Set these environment variables:")
            print(f'   export AZURE_CLIENT_ID="{CLIENT_ID}"')
            print(f'   export AZURE_CLIENT_SECRET="{CLIENT_SECRET}"')
            print("2. Save the MCP server script")
            print("3. Configure Claude Desktop")
            print("4. Restart Claude Desktop")
            
            return True
            
        else:
            print(f"❌ Token request failed: {response.status_code}")
            print(f"Error: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Error during token exchange: {e}")
        return False

def test_token(access_token):
    """Test the access token with a simple API call"""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        # Test user info
        response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
        if response.status_code == 200:
            user = response.json()
            print(f"✅ Connected as: {user.get('displayName', 'Unknown')} ({user.get('mail', user.get('userPrincipalName', 'No email'))})")
        else:
            print(f"⚠️  API test failed: {response.status_code}")
            
    except Exception as e:
        print(f"⚠️  API test error: {e}")

if __name__ == "__main__":
    authenticate()
