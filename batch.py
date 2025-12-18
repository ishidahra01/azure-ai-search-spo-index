from azure.identity import DefaultAzureCredential
import requests

credential = DefaultAzureCredential()
token = credential.get_token("https://graph.microsoft.com/.default")

headers = {
    "Authorization": f"Bearer {token.token}"
}

resp = requests.get(
    "https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root/children",
    headers=headers
)
print(resp.json())
