# Token Tormentor
Token Tormentor was designed to demonstrate the capabilities of Device Code Phishing in Entra ID. It uses the FOCI1 group to exchange refresh tokens for access tokens of different clients and resources. This allows interaction with the victim's Entra ID account in various ways.

The following features have been integrated

- Download all files from a user's default OneDrive drive
- Upload a file to a user's desktop folder in OneDrive
- Download recent Teams conversations via the Skype API
- Send Teams messages in recent conversations via the Skype API
- Download all of a user's emails as EML
- Send emails on behalf of the user
- Add Outlook forwarding rules
- Retrieve BitLocker recovery key from user-owned systems

Token Tormentor can also interact with roadTools (https://github.com/dirkjanm/ROADtools) and AzureHound (https://github.com/BloodHoundAD/AzureHound). In bove cases Token Tormentor handels the tokens for you.

## Installation
```
$ pip install -r requirements.txt
```

## Usage
TokenTormentor.py needs a file that contains the necessary tokens
```
TokenTormentor.py token
```

The required token JSON format is basically just the HTTP response of a completed OAuth 2.0 Device Authorisation Grant flow.:
```
{
    "access_token": "eyJ0eXAiOiJKV1Qi[cut]",
    "expires_in": 1199,
    "ext_expires_in": 1199,
    "foci": "1",
    "id_token": "eyJ0eXAiOiJKV1QiLCJhbGci[cut]",
    "refresh_token": "0.AXkA2CxcknehpUGT8xJX_[cut]",
    "scope": "email openid profile https://graph.microsoft.com/AuditLog.Read.All https://graph.microsoft.com/Calendar.ReadWrite [cut]",
    "token_type": "Bearer"
}
```
