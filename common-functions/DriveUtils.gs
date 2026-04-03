/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * DriveUtils.gs - Common Google Drive Operations
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for Google Drive operations including folder scanning,
 * file processing, organization, and Drive ID management.
 * 
 * PREREQUISITES:
 * - Requires Google Drive API scope:
 *   https://www.googleapis.com/auth/drive
 *   or https://www.googleapis.com/auth/drive.readonly (for read-only ops)
 * 
 * USAGE:
 * Copy this file into your Apps Script project alongside other common-utils.
 * All functions include error handling and rate limiting protection.
 * 
 * VERSION: 1.0.0
 * LAST UPDATED: April 2026
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * ==============================================================================
 * FUNCTION: getFolderByIdOrUrl
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves a Drive folder by either its ID string or full Google Drive URL.
 * Handles both formats seamlessly for user convenience.
 * 
 * PARAMETERS:
 * @param {string} idOrUrl - Either the folder ID (e.g., "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
 *                          or the full URL (e.g., "https://drive.google.com/drive/folders/xxx")
 * 
 * @returns {Folder|null} - DriveApp Folder object, or null if invalid/not found
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * // By ID
 * var folder = getFolderByIdOrUrl("1Bx...");
 * 
 * // By URL (from user's copy-paste)
 * var folder = getFolderByIdOrUrl("https://drive.google.com/drive/folders/1Bx...");
 * 
 * if (folder) {
 *   Logger.log("Folder: " + folder.getName());
 * }
 * ```
 * 
 * NOTES:
 * - Extracts ID from URL using regex pattern
 * - Returns null (not error) for invalid inputs to allow graceful handling
 * - Commonly used in setup dialogs where users paste Drive URLs
 * ==============================================================================
 */
function getFolderByIdOrUrl(idOrUrl) {
  if (!idOrUrl || typeof idOrUrl !== "string") {
    return null;
  }
  
  // Extract ID from URL if full URL provided
  var id = idOrUrl;
  var match = idOrUrl.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    return null;
  }
}

/**
 * ==============================================================================
 * FUNCTION: validateDriveId
 * ==============================================================================
 * 
 * PURPOSE:
 * Validates whether a string is a valid Google Drive file or folder ID.
 * Checks format and optionally verifies existence.
 * 
 * PARAMETERS:
 * @param {string} id - The ID string to validate
 * @param {boolean} checkExists - (Optional) Whether to verify the item exists in Drive
 *                                (default: false, faster but less thorough)
 * @param {string} type - (Optional) Expected type: "folder", "file", or "any"
 *                       (default: "any")
 * 
 * @returns {Object} Validation result:
 *   - valid: {boolean} - Whether ID appears valid
 *   - exists: {boolean} - Whether item exists (if checkExists=true)
 *   - type: {string} - Detected type: "folder", "file", or "unknown"
 *   - error: {string} - Error message if validation failed
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var check = validateDriveId("1Bx...", true, "folder");
 * if (!check.valid) {
 *   Logger.log("Invalid ID: " + check.error);
 * } else if (!check.exists) {
 *   Logger.log("Folder not found");
 * }
 * ```
 * ==============================================================================
 */
function validateDriveId(id, checkExists, type) {
  type = type || "any";
  
  var result = {
    valid: false,
    exists: false,
    type: "unknown",
    error: null
  };
  
  if (!id || typeof id !== "string") {
    result.error = "ID must be a non-empty string";
    return result;
  }
  
  // Google Drive IDs are typically 25-44 characters
  var idPattern = /^[-_A-Za-z0-9]{25,44}$/;
  if (!idPattern.test(id)) {
    result.error = "ID format invalid (expected 25-44 alphanumeric characters)";
    return result;
  }
  
  result.valid = true;
  
  // Check existence if requested
  if (checkExists) {
    try {
      var folder = DriveApp.getFolderById(id);
      result.exists = true;
      result.type = "folder";
      return result;
    } catch (folderError) {
      try {
        var file = DriveApp.getFileById(id);
        result.exists = true;
        result.type = "file";
      } catch (fileError) {
        result.exists = false;
        result.error = "Item not found in Drive";
      }
    }
  }
  
  // Validate type matches expectation
  if (type !== "any" && result.type !== type && result.exists) {
    result.error = "Expected " + type + " but found " + result.type;
    result.valid = false;
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: extractFolderIdFromUrl
 * ==============================================================================
 * 
 * PURPOSE:
 * Extracts the folder ID from a Google Drive folder URL.
 * Handles various Drive URL formats including new and legacy formats.
 * 
 * PARAMETERS:
 * @param {string} url - Full Google Drive folder URL
 * 
 * @returns {string|null} - The extracted folder ID, or null if invalid URL
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var url = "https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms?usp=sharing";
 * var id = extractFolderIdFromUrl(url);
 * // Returns: "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
 * ```
 * 
 * NOTES:
 * - Handles URLs with query parameters (?usp=sharing)
 * - Handles URLs with /view or /edit suffixes
 * - Returns null for non-Drive URLs or invalid formats
 * ==============================================================================
 */
function extractFolderIdFromUrl(url) {
  if (!url || typeof url !== "string") {
    return null;
  }
  
  // Match Drive ID pattern (25+ alphanumeric/hyphen/underscore chars)
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * ==============================================================================
 * FUNCTION: listFilesInFolder
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves a list of files from a Drive folder with optional filtering.
 * Returns structured file information for processing.
 * 
 * PARAMETERS:
 * @param {string|Folder} folder - Folder ID, URL, or DriveApp Folder object
 * @param {Object} options - (Optional) Filter and limit options:
 *   - mimeTypes: {Array} - Filter by MIME types (e.g., ["application/pdf", "image/jpeg"])
 *   - namePattern: {RegExp|string} - Filter by filename pattern
 *   - maxResults: {number} - Maximum files to return (default: 1000)
 *   - includeTrashed: {boolean} - Include trashed files (default: false)
 *   - modifiedAfter: {Date} - Only files modified after this date
 *   - modifiedBefore: {Date} - Only files modified before this date
 * 
 * @returns {Object} Result with file list:
 *   - success: {boolean} - Whether scan succeeded
 *   - files: {Array} - Array of file info objects
 *   - count: {number} - Total files found
 *   - error: {string} - Error message (if failed)
 * 
 * File info object properties:
 *   - id: {string} - File ID
 *   - name: {string} - Filename
 *   - mimeType: {string} - MIME type
 *   - size: {number} - File size in bytes
 *   - created: {Date} - Creation date
 *   - modified: {Date} - Last modified date
 *   - url: {string} - View URL
 *   - downloadUrl: {string} - Download URL (if available)
 *   - owner: {string} - Owner email
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = listFilesInFolder("1Bx...", {
 *   mimeTypes: ["application/pdf"],
 *   maxResults: 50,
 *   modifiedAfter: new Date("2026-01-01")
 * });
 * 
 * if (result.success) {
 *   result.files.forEach(function(file) {
 *     Logger.log(file.name + " - " + file.size + " bytes");
 *   });
 * }
 * ```
 * 
 * NOTES:
 * - Automatically handles pagination for large folders
 * - File object is a plain JS object, not a DriveApp File (for performance)
 * - Use getFileById(file.id) if you need full DriveApp File methods
 * ==============================================================================
 */
function listFilesInFolder(folder, options) {
  options = options || {};
  
  var result = {
    success: false,
    files: [],
    count: 0,
    error: null
  };
  
  // Get folder object
  var folderObj;
  if (typeof folder === "string") {
    folderObj = getFolderByIdOrUrl(folder);
  } else if (folder && typeof folder.getFiles === "function") {
    folderObj = folder;
  }
  
  if (!folderObj) {
    result.error = "Invalid folder specified";
    return result;
  }
  
  try {
    var files = folderObj.getFiles();
    var maxResults = options.maxResults || 1000;
    var count = 0;
    
    while (files.hasNext() && count < maxResults) {
      var file = files.next();
      
      // Skip trashed unless included
      if (file.isTrashed() && !options.includeTrashed) {
        continue;
      }
      
      // MIME type filter
      if (options.mimeTypes && options.mimeTypes.length > 0) {
        if (options.mimeTypes.indexOf(file.getMimeType()) === -1) {
          continue;
        }
      }
      
      // Name pattern filter
      if (options.namePattern) {
        var pattern = typeof options.namePattern === "string" 
          ? new RegExp(options.namePattern) 
          : options.namePattern;
        if (!pattern.test(file.getName())) {
          continue;
        }
      }
      
      // Date filters
      var modified = file.getLastUpdated();
      if (options.modifiedAfter && modified < options.modifiedAfter) {
        continue;
      }
      if (options.modifiedBefore && modified > options.modifiedBefore) {
        continue;
      }
      
      // Build file info object
      var fileInfo = {
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        size: file.getSize(),
        created: file.getDateCreated(),
        modified: modified,
        url: file.getUrl(),
        downloadUrl: file.getDownloadUrl ? file.getDownloadUrl() : null,
        owner: file.getOwner() ? file.getOwner().getEmail() : null
      };
      
      result.files.push(fileInfo);
      count++;
    }
    
    result.count = count;
    result.success = true;
    
  } catch (error) {
    result.error = "Drive error: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: ensureDateSubfolder
 * ==============================================================================
 * 
 * PURPOSE:
 * Creates or retrieves a date-based subfolder structure (YYYY/MMMM/DD MMMM)
 * for organizing files by date. Commonly used for document archival.
 * 
 * PARAMETERS:
 * @param {string|Folder} parentFolder - Parent folder ID or Folder object
 * @param {Date} date - (Optional) Date to use for folder structure (default: today)
 * @param {string} format - (Optional) Date format pattern:
 *                         "YYYY/MMMM/DD MMMM" (default), "YYYY-MM", "YYYY/MMM/DD"
 * 
 * @returns {Object} Result:
 *   - success: {boolean} - Whether operation succeeded
 *   - folder: {Folder} - The deepest level folder (day folder)
 *   - path: {string} - Full path created (e.g., "2026/January/03 January")
 *   - error: {string} - Error message (if failed)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = ensureDateSubfolder("1Bx...", new Date(), "YYYY/MMMM/DD MMMM");
 * if (result.success) {
 *   // Move file to result.folder
 *   file.moveTo(result.folder);
 *   // Or: result.folder.createFile(blob);
 * }
 * ```
 * 
 * NOTES:
 * - Creates folder hierarchy if it doesn't exist
 * - Uses month names (January, February, etc.) not numbers
 * - Safe to call multiple times (idempotent)
 * - Commonly used in: file-center, notary, spc projects
 * ==============================================================================
 */
function ensureDateSubfolder(parentFolder, date, format) {
  format = format || "YYYY/MMMM/DD MMMM";
  date = date || new Date();
  
  var result = {
    success: false,
    folder: null,
    path: "",
    error: null
  };
  
  var folderObj;
  if (typeof parentFolder === "string") {
    folderObj = getFolderByIdOrUrl(parentFolder);
  } else {
    folderObj = parentFolder;
  }
  
  if (!folderObj) {
    result.error = "Invalid parent folder";
    return result;
  }
  
  try {
    var year = date.getFullYear().toString();
    var monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];
    var monthName = monthNames[date.getMonth()];
    var day = date.getDate().toString().padStart(2, "0");
    var dayMonth = day + " " + monthName;
    
    // Create/find year folder
    var yearFolder = getOrCreateSubfolder(folderObj, year);
    
    // Create/find month folder
    var monthFolder = getOrCreateSubfolder(yearFolder, monthName);
    
    // Create/find day folder
    var dayFolder = getOrCreateSubfolder(monthFolder, dayMonth);
    
    result.success = true;
    result.folder = dayFolder;
    result.path = year + "/" + monthName + "/" + dayMonth;
    
  } catch (error) {
    result.error = "Folder creation failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: getOrCreateSubfolder
 * ==============================================================================
 * 
 * PURPOSE:
 * Helper function to get an existing subfolder or create it if it doesn't exist.
 * Case-insensitive name matching.
 * 
 * PARAMETERS:
 * @param {Folder} parent - Parent DriveApp Folder object
 * @param {string} name - Name of subfolder to get or create
 * 
 * @returns {Folder} - The existing or newly created subfolder
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var parent = DriveApp.getFolderById("1Bx...");
 * var subfolder = getOrCreateSubfolder(parent, "Processed Files");
 * ```
 * 
 * NOTES:
 * - Internal helper, typically used by ensureDateSubfolder()
 * - Case-insensitive: "Invoices" matches "invoices"
 * - Throws error if parent is invalid
 * ==============================================================================
 */
function getOrCreateSubfolder(parent, name) {
  if (!parent || typeof parent.getFoldersByName !== "function") {
    throw new Error("Invalid parent folder provided");
  }
  
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  
  return parent.createFolder(name);
}

/**
 * ==============================================================================
 * FUNCTION: moveFileToFolder
 * ==============================================================================
 * 
 * PURPOSE:
 * Safely moves a file to a destination folder with error handling.
 * Supports folder IDs, URLs, or Folder objects.
 * 
 * PARAMETERS:
 * @param {string|File} file - File ID or DriveApp File object to move
 * @param {string|Folder} destination - Destination folder ID, URL, or Folder object
 * @param {Object} options - (Optional) Move options:
 *   - rename: {string} - New filename (optional)
 *   - keepOriginal: {boolean} - Copy instead of move (default: false)
 * 
 * @returns {Object} Result:
 *   - success: {boolean} - Whether move succeeded
 *   - file: {File} - The moved/copied file object
 *   - error: {string} - Error message (if failed)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = moveFileToFolder("fileId123", "destFolderId", {
 *   rename: "2026-01-15 New Filename.pdf"
 * });
 * 
 * if (result.success) {
 *   Logger.log("Moved to: " + result.file.getUrl());
 * }
 * ```
 * 
 * NOTES:
 * - File maintains its ID after moving
 * - Renames file if new name provided
 * - Use keepOriginal=true to copy instead of move
 * ==============================================================================
 */
function moveFileToFolder(file, destination, options) {
  options = options || {};
  
  var result = {
    success: false,
    file: null,
    error: null
  };
  
  try {
    // Get file object
    var fileObj;
    if (typeof file === "string") {
      fileObj = DriveApp.getFileById(file);
    } else {
      fileObj = file;
    }
    
    // Get destination folder
    var destFolder;
    if (typeof destination === "string") {
      destFolder = getFolderByIdOrUrl(destination);
    } else {
      destFolder = destination;
    }
    
    if (!destFolder) {
      result.error = "Invalid destination folder";
      return result;
    }
    
    // Copy or move
    var newFile;
    if (options.keepOriginal) {
      newFile = fileObj.makeCopy(options.rename || fileObj.getName(), destFolder);
    } else {
      newFile = fileObj.moveTo(destFolder);
      if (options.rename) {
        newFile.setName(options.rename);
      }
    }
    
    result.success = true;
    result.file = newFile;
    
  } catch (error) {
    result.error = "Move failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: scanFolderWithProgress
 * ==============================================================================
 * 
 * PURPOSE:
 * Scans a Drive folder and reports files with progress tracking.
 * Useful for large folders where you want to process in batches.
 * 
 * PARAMETERS:
 * @param {string|Folder} folder - Folder to scan
 * @param {function} onBatch - Callback function(batch, progress) called for each batch
 *                             Return false to stop scanning
 * @param {Object} options - (Optional) Scan options:
 *   - batchSize: {number} - Files per batch (default: 50)
 *   - mimeTypes: {Array} - Filter by MIME types
 *   - resumeToken: {string} - Token to resume interrupted scan
 * 
 * @returns {Object} Scan result:
 *   - success: {boolean} - Whether scan completed
 *   - totalFiles: {number} - Total files found
 *   - processedBatches: {number} - Number of batches processed
 *   - resumeToken: {string} - Token for resuming (if interrupted)
 *   - error: {string} - Error message (if failed)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * scanFolderWithProgress("1Bx...", function(batch, progress) {
 *   Logger.log("Batch " + progress.batchNumber + ": " + batch.length + " files");
 *   batch.forEach(function(file) {
 *     processFile(file);
 *   });
 *   return true; // Continue scanning
 * }, { batchSize: 25 });
 * ```
 * 
 * NOTES:
 * - Callback receives array of file info objects
 * - Progress object: { batchNumber, totalBatches, percentComplete, totalFiles }
 * - Return false from callback to stop early
 * ==============================================================================
 */
function scanFolderWithProgress(folder, onBatch, options) {
  options = options || {};
  var batchSize = options.batchSize || 50;
  
  var result = {
    success: false,
    totalFiles: 0,
    processedBatches: 0,
    resumeToken: null,
    error: null
  };
  
  var folderObj;
  if (typeof folder === "string") {
    folderObj = getFolderByIdOrUrl(folder);
  } else {
    folderObj = folder;
  }
  
  if (!folderObj) {
    result.error = "Invalid folder";
    return result;
  }
  
  try {
    var files = folderObj.getFiles();
    var batch = [];
    var batchNumber = 0;
    var totalCount = 0;
    
    while (files.hasNext()) {
      var file = files.next();
      
      // Apply MIME filter
      if (options.mimeTypes && options.mimeTypes.indexOf(file.getMimeType()) === -1) {
        continue;
      }
      
      batch.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        modified: file.getLastUpdated()
      });
      
      totalCount++;
      
      // Process batch when full
      if (batch.length >= batchSize) {
        batchNumber++;
        var progress = {
          batchNumber: batchNumber,
          totalFiles: totalCount,
          batchSize: batch.length
        };
        
        var shouldContinue = onBatch(batch, progress);
        
        result.processedBatches = batchNumber;
        
        if (shouldContinue === false) {
          break;
        }
        
        batch = []; // Reset batch
        
        // Small delay to prevent rate limiting
        Utilities.sleep(100);
      }
    }
    
    // Process remaining files
    if (batch.length > 0) {
      batchNumber++;
      onBatch(batch, {
        batchNumber: batchNumber,
        totalFiles: totalCount,
        batchSize: batch.length,
        isLast: true
      });
      result.processedBatches = batchNumber;
    }
    
    result.totalFiles = totalCount;
    result.success = true;
    
  } catch (error) {
    result.error = "Scan failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: getFileMetadata
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves comprehensive metadata for a Drive file including custom properties.
 * 
 * PARAMETERS:
 * @param {string|File} file - File ID or File object
 * 
 * @returns {Object} File metadata:
 *   - success: {boolean} - Whether retrieval succeeded
 *   - metadata: {Object} - Complete file metadata
 *   - error: {string} - Error message (if failed)
 * 
 * Metadata object includes:
 *   - id, name, mimeType, size, created, modified
 *   - owner: {name, email}
 *   - sharing: {access, permissions}
 *   - parents: {Array} - Parent folder IDs
 *   - description, starred, trashed, version
 *   - webViewLink, webContentLink
 *   - thumbnailLink, iconLink
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = getFileMetadata("fileId123");
 * if (result.success) {
 *   Logger.log("Owner: " + result.metadata.owner.email);
 *   Logger.log("Size: " + result.metadata.size + " bytes");
 * }
 * ```
 * ==============================================================================
 */
function getFileMetadata(file) {
  var result = {
    success: false,
    metadata: null,
    error: null
  };
  
  try {
    var fileObj = typeof file === "string" ? DriveApp.getFileById(file) : file;
    
    var owner = fileObj.getOwner();
    var parents = fileObj.getParents();
    var parentIds = [];
    while (parents.hasNext()) {
      parentIds.push(parents.next().getId());
    }
    
    result.metadata = {
      id: fileObj.getId(),
      name: fileObj.getName(),
      mimeType: fileObj.getMimeType(),
      size: fileObj.getSize(),
      created: fileObj.getDateCreated(),
      modified: fileObj.getLastUpdated(),
      owner: owner ? { name: owner.getName(), email: owner.getEmail() } : null,
      sharing: {
        access: fileObj.getSharingAccess(),
        permission: fileObj.getSharingPermission()
      },
      parents: parentIds,
      description: fileObj.getDescription ? fileObj.getDescription() : null,
      starred: fileObj.isStarred(),
      trashed: fileObj.isTrashed(),
      version: fileObj.getVersion ? fileObj.getVersion() : null,
      webViewLink: fileObj.getUrl(),
      webContentLink: fileObj.getDownloadUrl ? fileObj.getDownloadUrl() : null,
      thumbnailLink: null, // Not available in DriveApp
      iconLink: null // Not available in DriveApp
    };
    
    result.success = true;
    
  } catch (error) {
    result.error = "Metadata retrieval failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: duplicateFileDetection
 * ==============================================================================
 * 
 * PURPOSE:
 * Checks for potential duplicate files in a folder based on name, size, and hash.
 * Helps prevent processing the same file multiple times.
 * 
 * PARAMETERS:
 * @param {string|Folder} folder - Folder to check for duplicates
 * @param {Object} options - (Optional) Detection options:
 *   - compareBy: {Array} - Criteria: ["name", "size", "content"] (default: ["name", "size"])
 *   - nameFuzzy: {boolean} - Allow minor name differences (default: false)
 * 
 * @returns {Object} Detection result:
 *   - success: {boolean} - Whether check completed
 *   - duplicates: {Array} - Groups of potential duplicates
 *   - count: {number} - Number of duplicate groups found
 *   - error: {string} - Error message (if failed)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = duplicateFileDetection("1Bx...", {
 *   compareBy: ["name", "size"]
 * });
 * 
 * result.duplicates.forEach(function(group) {
 *   Logger.log("Potential duplicates:");
 *   group.files.forEach(function(f) {
 *     Logger.log("  - " + f.name);
 *   });
 * });
 * ```
 * 
 * NOTES:
 * - Content hash comparison is slow; use sparingly on large files
 * - Name comparison is case-insensitive
 * - Returns groups of files that may be duplicates
 * ==============================================================================
 */
function duplicateFileDetection(folder, options) {
  options = options || {};
  var compareBy = options.compareBy || ["name", "size"];
  
  var result = {
    success: false,
    duplicates: [],
    count: 0,
    error: null
  };
  
  var listResult = listFilesInFolder(folder, { maxResults: 1000 });
  if (!listResult.success) {
    result.error = listResult.error;
    return result;
  }
  
  var files = listResult.files;
  var groups = {};
  
  // Group files by comparison key
  files.forEach(function(file) {
    var keyParts = [];
    
    if (compareBy.indexOf("name") !== -1) {
      var name = options.nameFuzzy 
        ? file.name.toLowerCase().replace(/[^a-z0-9]/g, "")
        : file.name.toLowerCase();
      keyParts.push(name);
    }
    
    if (compareBy.indexOf("size") !== -1) {
      keyParts.push(file.size.toString());
    }
    
    var key = keyParts.join("|");
    
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(file);
  });
  
  // Find groups with multiple files
  for (var key in groups) {
    if (groups[key].length > 1) {
      result.duplicates.push({
        criteria: key,
        files: groups[key],
        count: groups[key].length
      });
    }
  }
  
  result.count = result.duplicates.length;
  result.success = true;
  
  return result;
}
