/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'sentitems': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Display names (English and Swedish, lower case) of the default folders and
 * their Graph well-known folder names, which work in place of a folder id.
 * Lets "Inbox/Sub" resolve in a mailbox where the inbox is called "Inkorg".
 */
const WELL_KNOWN_FOLDER_NAMES = {
  'inbox': 'inbox',
  'inkorg': 'inbox',
  'sent': 'sentitems',
  'sentitems': 'sentitems',
  'sent items': 'sentitems',
  'skickat': 'sentitems',
  'skickade objekt': 'sentitems',
  'drafts': 'drafts',
  'utkast': 'drafts',
  'deleted': 'deleteditems',
  'deleteditems': 'deleteditems',
  'deleted items': 'deleteditems',
  'borttaget': 'deleteditems',
  'borttagna objekt': 'deleteditems',
  'archive': 'archive',
  'arkiv': 'archive',
  'junk': 'junkemail',
  'junkemail': 'junkemail',
  'junk email': 'junkemail',
  'skräppost': 'junkemail'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {

  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  try {
    // Try to find the folder by name
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (folderId) {
      const path = `me/mailFolders/${encodeURIComponent(folderId)}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, fall back to inbox
    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return WELL_KNOWN_FOLDERS['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return WELL_KNOWN_FOLDERS['inbox'];
  }
}

/**
 * Get the ID of a child folder by name within a specific parent folder.
 * @param {string} accessToken - Access token
 * @param {string|null} parentFolderId - Parent folder ID, or null to search root
 * @param {string} name - Folder display name to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getChildFolderIdByName(accessToken, parentFolderId, name) {
  try {
    const endpoint = parentFolderId
      ? `me/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`
      : 'me/mailFolders';

    // Try exact match first
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, {
      $filter: `displayName eq '${name.replace(/'/g, "''")}'`, // OData escapes ' as ''
      $select: 'id,displayName'
    });
    if (response.value && response.value.length > 0) {
      return response.value[0].id;
    }

    // Fallback: case-insensitive scan
    const allResponse = await callGraphAPI(accessToken, 'GET', endpoint, null, {
      $top: 100,
      $select: 'id,displayName'
    });
    if (allResponse.value) {
      const lower = name.toLowerCase();
      const match = allResponse.value.find(f => f.displayName.toLowerCase() === lower);
      if (match) return match.id;
    }
    return null;
  } catch (error) {
    console.error(`Error finding child folder "${name}": ${error.message}`);
    return null;
  }
}

/**
 * Get the real ID of a default folder from its Graph well-known name.
 * (Message moves accept well-known names, but e.g. rule actions need the ID.)
 * @param {string} accessToken - Access token
 * @param {string} wellKnownName - e.g. "inbox", "sentitems"
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getWellKnownFolderId(accessToken, wellKnownName) {
  try {
    const folder = await callGraphAPI(accessToken, 'GET', `me/mailFolders/${wellKnownName}`, null, { $select: 'id' });
    return (folder && folder.id) || null;
  } catch (error) {
    console.error(`Error finding well-known folder "${wellKnownName}": ${error.message}`);
    return null;
  }
}

/**
 * Get the ID of a mail folder by its name or path.
 * Supports path notation like "Inbox/2024/Viktigt" to resolve nested folders.
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name or slash-separated path
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  try {
    console.error(`Looking for folder "${folderName}"`);

    // Path notation: resolve each segment in turn
    if (folderName.includes('/')) {
      const parts = folderName.split('/').map(p => p.trim()).filter(Boolean);
      // A default folder as the first segment ("Inbox", "Inkorg", ...) is found by its well-known name
      const wellKnownRoot = WELL_KNOWN_FOLDER_NAMES[(parts[0] || '').toLowerCase()];
      let currentId = wellKnownRoot ? await getWellKnownFolderId(accessToken, wellKnownRoot) : null;
      if (wellKnownRoot && !currentId) {
        return null;
      }
      for (const part of wellKnownRoot ? parts.slice(1) : parts) {
        currentId = await getChildFolderIdByName(accessToken, currentId, part);
        if (!currentId) {
          console.error(`Path segment "${part}" not found`);
          return null;
        }
      }
      console.error(`Resolved path "${folderName}" to ID: ${currentId}`);
      return currentId;
    }

    // Simple name: a default folder, or search root-level folders
    const wellKnown = WELL_KNOWN_FOLDER_NAMES[folderName.trim().toLowerCase()];
    if (wellKnown) {
      return await getWellKnownFolderId(accessToken, wellKnown);
    }
    const id = await getChildFolderIdByName(accessToken, null, folderName);
    if (id) {
      console.error(`Found folder "${folderName}" with ID: ${id}`);
      return id;
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );
    
    if (!response.value) {
      return [];
    }
    
    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);
    
    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${encodeURIComponent(folder.id)}/childFolders`,
          null,
          { 
            $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
          }
        );
        
        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });
    
    const childFolders = await Promise.all(childFolderPromises);
    
    // Combine top-level folders and all child folders
    return [...response.value, ...childFolders.flat()];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  WELL_KNOWN_FOLDER_NAMES,
  resolveFolderPath,
  getFolderIdByName,
  getChildFolderIdByName,
  getAllFolders
};
