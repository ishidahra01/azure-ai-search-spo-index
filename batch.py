from azure.identity import DefaultAzureCredential
import requests

credential = DefaultAzureCredential()
token = credential.get_token("https://graph.microsoft.com/.default")
site_id = "mngenvmcap873995.sharepoint.com,b265cdd9-f76a-46c2-9919-291f7405a002,c4627576-656c-436f-b686-aeca7805b032"

headers = {
    "Authorization": f"Bearer {token.token}"
}

resp = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children",
    headers=headers
)
print(resp.json())
