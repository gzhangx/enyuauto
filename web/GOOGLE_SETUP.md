# Google OAuth Setup

## 1. Create Google OAuth Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Sheets API**:
   - Go to "APIs & Services" > "Library"
   - Search for "Google Sheets API"
   - Click "Enable"
4. Create OAuth 2.0 credentials:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Web application"
   - Add authorized JavaScript origins:
     - `http://localhost:5173` (for development)
     - Your production URL (when deploying)
   - Copy the **Client ID**

## 2. Configure the Application

1. Open `src/main.tsx`
2. Replace `YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com` with your actual Client ID

## 3. Set Up OAuth Consent Screen (if needed)

1. Go to "APIs & Services" > "OAuth consent screen"
2. Choose "External" user type
3. Fill in the required information
4. Add scopes:
   - `https://www.googleapis.com/auth/spreadsheets.readonly` (to read sheets)
5. Add test users (during development)

## 4. Run the Application

```bash
npm run dev
```

## Features

- **Google Login**: Users must authenticate with Google before accessing the app
- **Token Management**: Tokens are stored in localStorage until expiration
- **Auto Logout**: Automatically logs out when token expires
- **Projects Page**: Your existing project management interface
- **Google Sheets Reader**: Read data from any Google Sheet by ID (requires share permissions)

## Notes

- The token expires after 1 hour by default
- To read a Google Sheet, the sheet must be shared with your Google account or set to "Anyone with the link can view"
- The Sheet ID can be found in the URL: `https://docs.google.com/spreadsheets/d/SHEET_ID/edit`
