# SharePoint App-Based Authentication Setup Guide

## Step-by-Step Instructions

### Step 1: Register App in Azure AD

1. **Go to Azure Portal**
   - Visit: https://portal.azure.com
   - Sign in with your organization account

2. **Navigate to App Registrations**
   - Search for "Azure Active Directory" in the top search bar
   - Click on **Azure Active Directory** (or **Microsoft Entra ID**)
   - In the left menu, click **App registrations**
   - Click **+ New registration**

3. **Register the App**
   - **Name**: `Maturity Assessment Automation` (or any name you prefer)
   - **Supported account types**: Select **Accounts in this organizational directory only**
   - **Redirect URI**: Leave blank (not needed for this use case)
   - Click **Register**

4. **Copy the Application (Client) ID** 
   - After registration, you'll see the app overview page
   - **Copy the "Application (client) ID"** - you'll need this!
   - Save it somewhere safe

### Step 2: Create Client Secret

1. **Go to Certificates & Secrets**
   - In the left menu, click **Certificates & secrets**
   - Click **+ New client secret**

2. **Create the Secret**
   - **Description**: `Maturity Assessment Script`
   - **Expires**: Select **24 months** (or your preference)
   - Click **Add**

3. **Copy the Secret Value**
   - ⚠️ **IMPORTANT**: Copy the **Value** (not the Secret ID)
   - You'll only see it once! Save it immediately.
   - It looks like: `abc123~XYZ789...`

### Step 3: Grant SharePoint Permissions

1. **Go to API Permissions**
   - In the left menu, click **API permissions**
   - Click **+ Add a permission**

2. **Select SharePoint**
   - Click **Microsoft Graph** (or search for "SharePoint")
   - ⚠️ **IMPORTANT**: Select **Application permissions** (NOT Delegated)
   - This is required for app-based authentication

3. **Add Required Permissions**
   - Search for: `Sites.ReadWrite.All`
   - Check the box next to it
   - Click **Add permissions**

4. **Grant Admin Consent** (if you have permission)
   - Click **Grant admin consent for [Your Organization]**
   - Click **Yes** to confirm
   - Status should show "✓ Granted for [Your Organization]"

   **Note**: If you don't have admin rights, ask your IT admin to grant consent.

### Step 4: Set Environment Variables

Once you have:
- ✅ Application (Client) ID
- ✅ Client Secret Value

Run these commands in your terminal:

```bash
export SHAREPOINT_UPLOAD=true
export SHAREPOINT_AUTH_METHOD=app
export SHAREPOINT_CLIENT_ID='paste-your-client-id-here'
export SHAREPOINT_CLIENT_SECRET='paste-your-secret-value-here'
```

### Step 5: Make It Persistent (for automation)

Add to your `~/.zshrc`:

```bash
echo "" >> ~/.zshrc
echo "# SharePoint App Authentication" >> ~/.zshrc
echo "export SHAREPOINT_UPLOAD=true" >> ~/.zshrc
echo "export SHAREPOINT_AUTH_METHOD=app" >> ~/.zshrc
echo "export SHAREPOINT_CLIENT_ID='your-client-id'" >> ~/.zshrc
echo "export SHAREPOINT_CLIENT_SECRET='your-client-secret'" >> ~/.zshrc
source ~/.zshrc
```

### Step 6: Test the Connection

Run the script to test:

```bash
python maturity_assessment.py
```

You should see:
```
✓ App-based authentication successful
```

## Troubleshooting

### "Insufficient privileges" error
- Make sure admin consent was granted for `Sites.ReadWrite.All`
- Contact your IT admin if you don't have permission

### "Invalid client" error
- Double-check your Client ID and Secret are correct
- Make sure you copied the Secret **Value**, not the Secret ID

### "403 Forbidden" or "Access denied" error
- ⚠️ **Most common issue**: Make sure you used **Application permissions** (not Delegated)
- Verify the app has `Sites.ReadWrite.All` as an **Application permission**
- Check that admin consent was granted
- The app may need explicit access to the SharePoint site (contact IT admin)

## Security Notes

- ⚠️ Never commit your Client Secret to Git
- ✅ Keep it in environment variables only
- ✅ The secret is already in `.gitignore`
- ✅ Use long-lived secrets (24 months) for automation

