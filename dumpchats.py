from flask import Flask, redirect, url_for, session, request
from msal import ConfidentialClientApplication
import requests
import json
import os

app = Flask(__name__)
app.secret_key = "your_secret_key"  # Replace with your own secret key

# Azure AD configuration
CLIENT_ID = { CLIENT_ID }
CLIENT_SECRET = { CLIENT_SECRET }
TENANT_ID = { TENANT_ID}
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:5000/callback"
SCOPE = ["https://graph.microsoft.com/.default"]

# MSAL client
msal_app = ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)


@app.route("/")
def home():
    return '<a href="/login">Log in with Microsoft</a>'


@app.route("/login")
def login():
    # Redirect to the Microsoft login page
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)


@app.route("/callback")
def callback():
    # Handle the redirect from Microsoft login
    code = request.args.get("code")
    if not code:
        return "Error: No authorization code provided."

    # Exchange the authorization code for a token
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
    if "access_token" in result:
        # Save the token in the session
        session["access_token"] = result["access_token"]
        return redirect(url_for("chats"))
    else:
        return f"Error: {result.get('error_description', 'Unknown error')}"


@app.route("/profile")
def profile():
    # Access Microsoft Graph API using the token
    token = session.get("access_token")
    if not token:
        return redirect(url_for("login"))

    headers = {"Authorization": f"Bearer {token}"}
    graph_url = "https://graph.microsoft.com/v1.0/me"
    response = requests.get(graph_url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        return f"Error: {response.status_code}, {response.text}"

@app.route("/chats")
def chats():
    # Access Microsoft Graph API using the token
    token = session.get("access_token")
    if not token:
        return redirect(url_for("login"))

    chats = get_chats(token)

    for chat in chats:
        get_messages(token, chat)

    return [chat["topic"] for chat in chats]


def get_chats(token):
    headers = {"Authorization": f"Bearer {token}"}
    graph_url = "https://graph.microsoft.com/v1.0/me/chats/"

    chats = []

    while True:
        response = requests.get(graph_url, headers=headers)

        if response.status_code == 200:
            data, next = parse_chats(response.json())
            print(next)
            chats.extend(data)
            if next:
                graph_url = next
            else:
                break
        else:
            return f"Error: {response.status_code}, {response.text}"

    return chats


def parse_chats(data):
    chats = data.get("value", [])
    next = data.get("@odata.nextLink")

    chats_parsed=[]

    for chat in chats:
        chat_parsed = {
            "chat_id" : chat.get("id"),
            "chat_type" : chat.get("chatType"),
            "topic" : chat.get("topic")
        }

        chats_parsed.append(chat_parsed)

    return chats_parsed, next

def get_messages(token, chat):
    headers = {"Authorization": f"Bearer {token}"}
    graph_url = f"https://graph.microsoft.com/v1.0/me/chats/{chat['chat_id']}/messages"

    messages = []

    while True:
        response = requests.get(graph_url, headers=headers)

        if response.status_code == 200:
            data, next = parse_messages(response.json())
            print(next)
            messages.extend(data)
            if next:
                graph_url = next
            else:
                break
        else:
            return f"Error: {response.status_code}, {response.text}"

    with open(f"dump/{chat['chat_id']}.json", "w") as f:
        f.write(json.dumps(messages))

    return len(messages)


def parse_messages(data):
    messages = data.get("value", [])
    next = data.get("@odata.nextLink")

    messages_parsed=[]

    for message in messages:
        message_parsed = {
            "message_id" : message.get("id"),
            "from" : message.get("from"),
            "body" : message.get("body")
        }

        messages_parsed.append(message_parsed)

    return messages_parsed, next


if __name__ == "__main__":
    try:
        os.mkdir('dump')
    except FileExistsError:
        pass

    app.run(debug=True)