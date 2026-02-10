// Google Drive Service - Uses access token from Better Auth session

const AUTH_SERVER =
  process.env.NEXT_PUBLIC_BETTER_AUTH_URL ||
  (typeof window !== "undefined" ? window.location.origin : "");

// Cached access token with expiry
let cachedAccessToken: string | null = null;
let cachedTokenExpiry: number = 0;
const TOKEN_CACHE_MS = 50 * 60 * 1000; // 50 minutes (tokens last ~60 min)

// Fetch the Google access token from the server (uses session cookie)
export const getGoogleAccessToken = async (forceRefresh = false): Promise<string> => {
    if (!forceRefresh && cachedAccessToken && Date.now() < cachedTokenExpiry) {
        return cachedAccessToken;
    }
    
    const response = await fetch(`${AUTH_SERVER}/api/google-token`, {
        credentials: 'include',
    });
    
    if (!response.ok) {
        cachedAccessToken = null;
        cachedTokenExpiry = 0;
        const error = await response.json().catch(() => ({}));
        throw new Error(error.error || 'Failed to get Google token');
    }
    
    const data = await response.json();
    cachedAccessToken = data.accessToken;
    cachedTokenExpiry = Date.now() + TOKEN_CACHE_MS;
    return data.accessToken;
};

// Clear the cached token (call on logout)
export const clearGoogleToken = () => {
    cachedAccessToken = null;
    cachedTokenExpiry = 0;
};

// Helper: fetch with auto-retry on 401 (refreshes token and retries once)
const driveApiFetch = async (url: string, init?: RequestInit): Promise<Response> => {
    const accessToken = await getGoogleAccessToken();
    const headers = { ...init?.headers, 'Authorization': `Bearer ${accessToken}` };
    const response = await fetch(url, { ...init, headers });
    
    if (response.status === 401) {
        const freshToken = await getGoogleAccessToken(true);
        const retryHeaders = { ...init?.headers, 'Authorization': `Bearer ${freshToken}` };
        return fetch(url, { ...init, headers: retryHeaders });
    }
    
    return response;
};

// Check if user has Google Drive access
export const checkDriveAccess = async (): Promise<boolean> => {
    try {
        await getGoogleAccessToken();
        return true;
    } catch {
        return false;
    }
};

export const getFileContentAsBase64 = async (fileId: string): Promise<{base64: string, mimeType: string}> => {
    const response = await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);

    if (!response.ok) throw new Error(`Download failed: ${response.statusText}`);

    const blob = await response.blob();
    const mimeType = blob.type;

    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = reader.result as string;
            resolve({ base64: result.split(',')[1], mimeType });
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
};

export const extractFolderIdFromLink = (link: string): string | null => {
    const regex = /(?:\/folders\/|\/d\/|[?&]id=)([a-zA-Z0-9_-]{15,})/;
    const match = link.match(regex);
    return match ? match[1] : null;
};

export const isFolderLink = (link: string): boolean => {
    return /\/folders\//.test(link);
};

export const extractFileIdFromLink = (link: string): string | null => {
    // Matches: /file/d/ID, /document/d/ID, /spreadsheets/d/ID, /presentation/d/ID, ?id=ID
    const regex = /(?:\/(?:file|document|spreadsheets|presentation)\/d\/|[?&]id=)([a-zA-Z0-9_-]{15,})/;
    const match = link.match(regex);
    return match ? match[1] : null;
};

export const getFileMetadata = async (fileId: string): Promise<{ id: string; name: string; mimeType: string }> => {
    const url = `https://www.googleapis.com/drive/v3/files/${fileId}?fields=id,name,mimeType&supportsAllDrives=true`;
    const response = await driveApiFetch(url);
    if (!response.ok) {
        throw new Error(`Failed to get file metadata: ${response.status} ${response.statusText}`);
    }
    return response.json();
};

export const listFilesInFolder = async (folderId: string): Promise<Array<{name: string, id: string, mimeType: string}>> => {
    try {
        const query = `'${folderId}' in parents and trashed = false`;
        const allFiles: Array<{ name: string; id: string; mimeType: string }> = [];
        let pageToken: string | undefined;

        do {
            const url = new URL('https://www.googleapis.com/drive/v3/files');
            url.searchParams.append('pageSize', '1000');
            url.searchParams.append('fields', 'nextPageToken, files(id, name, mimeType)');
            url.searchParams.append('q', query);
            url.searchParams.append('supportsAllDrives', 'true');
            url.searchParams.append('includeItemsFromAllDrives', 'true');
            if (pageToken) url.searchParams.append('pageToken', pageToken);

            const response = await driveApiFetch(url.toString(), {
                method: 'GET',
                headers: { 'Content-Type': 'application/json' }
            });

            if (!response.ok) {
                let errorDetails = '';
                try {
                    const errorJson = await response.json();
                    errorDetails = errorJson?.error?.message ? `: ${errorJson.error.message}` : '';
                } catch {
                    // ignore JSON parsing errors
                }
                throw new Error(`Drive API Error: ${response.status} ${response.statusText}${errorDetails}`);
            }

            const data = await response.json();
            allFiles.push(...(data.files || []));
            pageToken = data.nextPageToken;
        } while (pageToken);

        return allFiles;

    } catch (err: any) {
        console.error("Error listing files", err);
        throw new Error(err.message || "Failed to list files");
    }
};

// List files in user's Drive root or search
export const listDriveFiles = async (query?: string): Promise<Array<{name: string, id: string, mimeType: string}>> => {
    const url = new URL('https://www.googleapis.com/drive/v3/files');
    url.searchParams.append('pageSize', '50');
    url.searchParams.append('fields', 'files(id, name, mimeType, modifiedTime)');
    url.searchParams.append('orderBy', 'modifiedTime desc');
    
    if (query) {
        url.searchParams.append('q', `name contains '${query}' and trashed = false`);
    } else {
        url.searchParams.append('q', 'trashed = false');
    }

    const response = await driveApiFetch(url.toString());

    if (!response.ok) {
        throw new Error(`Drive API Error: ${response.status}`);
    }

    const data = await response.json();
    return data.files || [];
};

// List folders in a specific parent folder (or root if no parentId)
export const listDriveFolders = async (
    parentId?: string,
    searchQuery?: string,
): Promise<Array<{ id: string; name: string }>> => {
    const url = new URL('https://www.googleapis.com/drive/v3/files');
    url.searchParams.append('pageSize', '100');
    url.searchParams.append('fields', 'files(id, name)');
    url.searchParams.append('orderBy', 'name');
    url.searchParams.append('supportsAllDrives', 'true');
    url.searchParams.append('includeItemsFromAllDrives', 'true');

    const conditions = ["mimeType = 'application/vnd.google-apps.folder'", "trashed = false"];
    if (parentId) {
        conditions.push(`'${parentId}' in parents`);
    }
    if (searchQuery) {
        conditions.push(`name contains '${searchQuery}'`);
    }
    url.searchParams.append('q', conditions.join(' and '));

    const response = await driveApiFetch(url.toString());
    if (!response.ok) {
        throw new Error(`Drive API Error: ${response.status}`);
    }
    const data = await response.json();
    return data.files || [];
};

// Upload a file (Blob) to Google Drive in the specified folder
export const uploadFileToDrive = async (
    blob: Blob,
    fileName: string,
    folderId: string,
    mimeType?: string,
): Promise<{ id: string; name: string; webViewLink: string }> => {
    const accessToken = await getGoogleAccessToken();

    const metadata = {
        name: fileName,
        parents: [folderId],
    };

    const form = new FormData();
    form.append(
        'metadata',
        new Blob([JSON.stringify(metadata)], { type: 'application/json' }),
    );
    form.append('file', blob, fileName);

    const response = await fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink&supportsAllDrives=true',
        {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}` },
            body: form,
        },
    );

    if (response.status === 401) {
        const freshToken = await getGoogleAccessToken(true);
        const retryResponse = await fetch(
            'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink&supportsAllDrives=true',
            {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${freshToken}` },
                body: form,
            },
        );
        if (!retryResponse.ok) {
            const err = await retryResponse.json().catch(() => ({}));
            throw new Error(err?.error?.message || `Upload failed: ${retryResponse.status}`);
        }
        return retryResponse.json();
    }

    if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err?.error?.message || `Upload failed: ${response.status}`);
    }

    return response.json();
};
