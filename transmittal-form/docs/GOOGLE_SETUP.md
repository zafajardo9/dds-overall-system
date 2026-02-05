# Google Cloud Console Setup Guide

This app uses two separate Google auth flows:
1. **Better Auth** - For user sign-in (server-side OAuth)
2. **Google Identity Services (GIS)** - For Google Drive file access (client-side)

## Required Configuration in Google Cloud Console

### Step 1: Configure OAuth Consent Screen
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Navigate to **APIs & Services** > **OAuth consent screen**
3. Set User Type to **External** (or Internal if using Google Workspace)
4. Fill in required fields (App name, support email, etc.)
5. Add scopes:
   - `https://www.googleapis.com/auth/drive.readonly`
   - `https://www.googleapis.com/auth/drive.metadata.readonly`
   - `openid`
   - `email`
   - `profile`

### Step 2: Configure OAuth Client ID (Credentials)
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **OAuth 2.0 Client IDs**
3. Application type: **Web application**
4. Add these **Authorized JavaScript origins**:
   ```
   http://localhost:3000
   http://localhost:8000
   ```
5. Add these **Authorized redirect URIs**:
   ```
   http://localhost:8000/api/auth/callback/google
   http://localhost:3000
   ```
6. Copy the **Client ID** and **Client Secret**

### Step 3: Create API Key (for Drive Picker)
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **API Key**
3. Click on the created key to configure restrictions:
   - **Application restrictions**: HTTP referrers
   - Add referrers:
     ```
     http://localhost:3000/*
     http://localhost:8000/*
     ```
   - **API restrictions**: Restrict to:
     - Google Drive API
     - Google Picker API
4. Copy the **API Key**

### Step 4: Enable Required APIs
1. Go to **APIs & Services** > **Library**
2. Search and enable:
   - **Google Drive API**
   - **Google Picker API**

## Update Your .env File

```env
BETTER_AUTH_SECRET=your-secret-key
BETTER_AUTH_URL=http://localhost:8000
VITE_BETTER_AUTH_URL=http://localhost:8000

GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=GOCSPX-your-client-secret
VITE_GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com

# API Key for Drive Picker (from Step 3)
VITE_GOOGLE_API_KEY=AIzaSy...your-api-key
```

## Troubleshooting

### "redirect_uri_mismatch" Error
- Ensure `http://localhost:8000/api/auth/callback/google` is in your OAuth redirect URIs
- Wait 5 minutes for changes to propagate

### "Access blocked" Error  
- Check OAuth consent screen is configured
- Ensure your email is added as a test user (if app is in testing mode)

### Drive Picker Not Working
- Verify `VITE_GOOGLE_API_KEY` is set in `.env`
- Check API Key restrictions allow localhost origins
- Ensure Google Drive API and Picker API are enabled

### CORS Errors
- Verify JavaScript origins include both `http://localhost:3000` and `http://localhost:8000`
