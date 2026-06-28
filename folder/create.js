/**
 * Create folder functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName, getChildFolderIdByName } = require('../email/folder-utils');

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const folderName = args.name;
  const parentFolder = args.parentFolder || '';
  
  if (!folderName) {
    return {
      content: [{ 
        type: "text", 
        text: "Folder name is required."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Create folder with appropriate parent
    const result = await createMailFolder(accessToken, folderName, parentFolder);
    
    return {
      content: [{ 
        type: "text", 
        text: result.message
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error creating folder: ${error.message}`
      }]
    };
  }
}

/**
 * Create a new mail folder
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to create
 * @param {string} parentFolderName - Name of the parent folder (optional)
 * @returns {Promise<object>} - Result object with status and message
 */
async function createMailFolder(accessToken, folderName, parentFolderName) {
  try {
    // Resolve parent folder first so we can scope the duplicate check correctly
    let parentId = null;
    let endpoint = 'me/mailFolders';
    if (parentFolderName) {
      parentId = await getFolderIdByName(accessToken, parentFolderName);
      if (!parentId) {
        return {
          success: false,
          message: `Parent folder "${parentFolderName}" not found. Please specify a valid parent folder or leave it blank to create at the root level.`
        };
      }
      endpoint = `me/mailFolders/${parentId}/childFolders`;
    }

    // Check for duplicate only within the target parent (not globally)
    const existingFolder = await getChildFolderIdByName(accessToken, parentId, folderName);
    if (existingFolder) {
      return {
        success: false,
        message: `A folder named "${folderName}" already exists${parentFolderName ? ` inside "${parentFolderName}"` : ' at the root level'}.`
      };
    }
    
    // Create the folder
    const folderData = {
      displayName: folderName
    };
    
    const response = await callGraphAPI(
      accessToken,
      'POST',
      endpoint,
      folderData
    );
    
    if (response && response.id) {
      const locationInfo = parentFolderName 
        ? `inside "${parentFolderName}"` 
        : "at the root level";
        
      return {
        success: true,
        message: `Successfully created folder "${folderName}" ${locationInfo}.`,
        folderId: response.id
      };
    } else {
      return {
        success: false,
        message: "Failed to create folder. The server didn't return a folder ID."
      };
    }
  } catch (error) {
    console.error(`Error creating folder "${folderName}": ${error.message}`);
    throw error;
  }
}

module.exports = handleCreateFolder;
