import { authenticate } from "@google-cloud/local-auth";
import { google } from "googleapis";
import fs from "fs";
import path from "path";

export const SCOPES = [
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
];

// Get credentials directory from environment variable or use default
const CREDS_DIR =
  process.env.GDRIVE_CREDS_DIR ||
  path.join(path.dirname(new URL(import.meta.url).pathname), "../../../");

// Ensure the credentials directory exists
function ensureCredsDirectory() {
  try {
    fs.mkdirSync(CREDS_DIR, { recursive: true });
    console.error(`Ensured credentials directory exists at: ${CREDS_DIR}`);
  } catch (error) {
    console.error(
      `Failed to create credentials directory: ${CREDS_DIR}`,
      error,
    );
    throw error;
  }
}

const credentialsPath = path.join(CREDS_DIR, ".gdrive-server-credentials.json");

async function authenticateWithTimeout(
  keyfilePath: string,
  SCOPES: string[],
  timeoutMs = 30000,
): Promise<any | null> {
  const timeoutPromise = new Promise((_, reject) =>
    setTimeout(() => reject(new Error("Authentication timed out")), timeoutMs),
  );

  const authPromise = authenticate({
    keyfilePath,
    scopes: SCOPES,
  });

  try {
    return await Promise.race([authPromise, timeoutPromise]);
  } catch (error) {
    console.error(error);
    return null;
  }
}

async function authenticateAndSaveCredentials() {
  console.error("Launching auth flow…");
  console.error("Using credentials path:", credentialsPath);

  const keyfilePath = path.join(CREDS_DIR, "gcp-oauth.keys.json");
  console.error("Using keyfile path:", keyfilePath);

  const auth = await authenticateWithTimeout(keyfilePath, SCOPES);
  if (!auth) {
    throw new Error("Authentication failed");
  }

  console.error("Authentication successful, saving credentials...");
  
  // Ensure directory exists before saving
  ensureCredsDirectory();

  // Save the credentials directly from the auth object
  fs.writeFileSync(credentialsPath, JSON.stringify(auth.credentials, null, 2));
  console.error("Credentials saved successfully to:", credentialsPath);
  
  return auth;
}

// Try to load credentials without prompting for auth
export async function loadCredentialsQuietly() {
  console.error("Attempting to load credentials from:", credentialsPath);

  if (!fs.existsSync(credentialsPath)) {
    console.error("No credentials file found");
    return null;
  }

  try {
    const savedCreds = JSON.parse(fs.readFileSync(credentialsPath, "utf-8"));
    console.error("Loaded existing credentials");

    const oauth2Client = new google.auth.OAuth2(
      savedCreds.client_id,
      savedCreds.client_secret,
      savedCreds.redirect_uris?.[0]
    );
    
    oauth2Client.setCredentials(savedCreds);

    // Check if token needs refresh
    const now = Date.now();
    const expiryTime = savedCreds.expiry_date || 0;
    
    if (expiryTime - now < 5 * 60 * 1000 && savedCreds.refresh_token) {
      console.error("Token expires soon, refreshing...");
      try {
        const { credentials } = await oauth2Client.refreshAccessToken();
        console.error("Token refreshed successfully");
        
        // Save the refreshed credentials
        ensureCredsDirectory();
        fs.writeFileSync(credentialsPath, JSON.stringify(credentials, null, 2));
        oauth2Client.setCredentials(credentials);
      } catch (error) {
        console.error("Failed to refresh token:", error);
        return null;
      }
    }

    return oauth2Client;
  } catch (error) {
    console.error("Error loading credentials:", error);
    return null;
  }
}

// Get valid credentials, prompting for auth if necessary
export async function getValidCredentials(forceAuth = false) {
  if (!forceAuth) {
    const quietAuth = await loadCredentialsQuietly();
    if (quietAuth) {
      return quietAuth;
    }
  }

  return await authenticateAndSaveCredentials();
}

// Background refresh that never prompts for auth
export function setupTokenRefresh() {
  console.error("Setting up automatic token refresh interval (45 minutes)");
  return setInterval(
    async () => {
      try {
        console.error("Running scheduled token refresh check");
        const auth = await loadCredentialsQuietly();
        if (auth) {
          console.error("Completed scheduled token refresh check");
        } else {
          console.error("Skipping token refresh - no valid credentials");
        }
      } catch (error) {
        console.error("Error in automatic token refresh:", error);
      }
    },
    45 * 60 * 1000,
  );
}
