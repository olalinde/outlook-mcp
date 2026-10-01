/**
 * Email attachments: building Graph file attachments from local files
 * and downloading a message's file attachments to disk.
 */
const fs = require('fs');
const path = require('path');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

// Graph accepts attachments up to 3 MB in a single request (larger ones need an upload session).
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const CONTENT_TYPES = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.pdf': 'application/pdf'
};

/**
 * Default content type from a file name's extension.
 * @param {string} fileName
 * @returns {string}
 */
function contentTypeFor(fileName) {
  return CONTENT_TYPES[path.extname(fileName || '').toLowerCase()] || 'application/octet-stream';
}

/**
 * Builds Graph `#microsoft.graph.fileAttachment` objects from local files.
 * @param {Array<{path:string, name?:string, contentType?:string, isInline?:boolean, contentId?:string}>} attachments
 * @returns {Array<object>} - Graph attachment objects
 * @throws {Error} - If a path is missing, not absolute, unreadable or larger than 3 MB
 */
function buildFileAttachments(attachments) {
  if (attachments === undefined || attachments === null) return [];
  if (!Array.isArray(attachments)) {
    throw new Error('attachments must be an array.');
  }

  return attachments.map((attachment) => {
    const filePath = attachment && attachment.path;
    if (!filePath || typeof filePath !== 'string') {
      throw new Error('Each attachment needs a "path".');
    }
    if (!path.isAbsolute(filePath)) {
      throw new Error(`Attachment path must be absolute: ${filePath}`);
    }

    let stat;
    try {
      stat = fs.statSync(filePath);
    } catch (error) {
      throw new Error(`Attachment file not found: ${filePath}`);
    }
    if (!stat.isFile()) {
      throw new Error(`Attachment path is not a file: ${filePath}`);
    }
    if (stat.size > MAX_ATTACHMENT_BYTES) {
      const mb = (stat.size / (1024 * 1024)).toFixed(1);
      throw new Error(`Attachment ${path.basename(filePath)} is ${mb} MB; the maximum is 3 MB per file.`);
    }

    const name = attachment.name || path.basename(filePath);
    const graphAttachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentType: attachment.contentType || contentTypeFor(name),
      contentBytes: fs.readFileSync(filePath).toString('base64'),
      isInline: attachment.isInline === true
    };
    if (attachment.contentId) {
      graphAttachment.contentId = attachment.contentId;
    }
    return graphAttachment;
  });
}

/**
 * Makes an attachment name safe to use as a file name in a single directory.
 * @param {string} name
 * @returns {string}
 */
function safeFileName(name) {
  const cleaned = String(name || '')
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, '_')
    .replace(/^[\s.]+|[\s.]+$/g, '');
  return cleaned || 'attachment';
}

/**
 * Returns a file name not already used in this batch ("a.png", "a-1.png", ...).
 * Existing files on disk with the same name are overwritten by design.
 */
function uniqueName(fileName, used) {
  let candidate = fileName;
  const ext = path.extname(fileName);
  const base = fileName.slice(0, fileName.length - ext.length);
  for (let i = 1; used.has(candidate.toLowerCase()); i++) {
    candidate = `${base}-${i}${ext}`;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

/**
 * download-attachments handler: saves a message's file attachments into saveDir.
 * @param {object} args - { messageId, saveDir, inlineOnly }
 * @returns {object} - MCP response with JSON text [{ path, name, contentType, contentId, isInline, size }]
 */
async function handleDownloadAttachments(args) {
  const { messageId, saveDir, inlineOnly = false } = args || {};

  if (!messageId) {
    return { content: [{ type: "text", text: "messageId is required." }] };
  }
  if (!saveDir || !path.isAbsolute(saveDir)) {
    return { content: [{ type: "text", text: "saveDir is required and must be an absolute path." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(accessToken, 'GET', `me/messages/${encodeURIComponent(messageId)}/attachments`);

    const fileAttachments = (response.value || [])
      .filter(a => a['@odata.type'] === '#microsoft.graph.fileAttachment')
      .filter(a => !inlineOnly || a.isInline === true);

    fs.mkdirSync(saveDir, { recursive: true });

    const used = new Set();
    const saved = [];
    for (const attachment of fileAttachments) {
      let contentBytes = attachment.contentBytes;
      if (!contentBytes) {
        const full = await callGraphAPI(accessToken, 'GET', `me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`);
        contentBytes = full.contentBytes || '';
      }
      const data = Buffer.from(contentBytes, 'base64');
      const filePath = path.join(saveDir, uniqueName(safeFileName(attachment.name), used));
      fs.writeFileSync(filePath, data);
      saved.push({
        path: filePath,
        name: attachment.name,
        contentType: attachment.contentType || contentTypeFor(attachment.name),
        contentId: attachment.contentId || null,
        isInline: attachment.isInline === true,
        size: data.length
      });
    }

    return { content: [{ type: "text", text: JSON.stringify(saved) }] };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    return { content: [{ type: "text", text: `Error downloading attachments: ${error.message}` }] };
  }
}

/**
 * Lists a message's attachments (metadata only, no content).
 * @returns {Promise<Array<{id, name, contentType, contentId, isInline, size}>>}
 */
async function listAttachmentMetadata(accessToken, messageId) {
  const response = await callGraphAPI(accessToken, 'GET', `me/messages/${encodeURIComponent(messageId)}/attachments`);
  return (response.value || []).map(a => ({
    id: a.id,
    name: a.name,
    contentType: a.contentType || null,
    contentId: a.contentId || null,
    isInline: a.isInline === true,
    size: a.size
  }));
}

module.exports = {
  MAX_ATTACHMENT_BYTES,
  contentTypeFor,
  buildFileAttachments,
  safeFileName,
  handleDownloadAttachments,
  listAttachmentMetadata
};
