/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const handleMarkAsRead = require('./mark-as-read');
const handleDeleteEmail = require('./delete');
const { handleDownloadAttachments } = require('./attachments');

// Shared schema for the attachments parameter of send-email and draft-email
const attachmentsSchema = {
  type: "array",
  description: "Files to attach. Each file max 3 MB. Use isInline + contentId and reference it in the HTML body as <img src=\"cid:<contentId>\">.",
  items: {
    type: "object",
    properties: {
      path: { type: "string", description: "Absolute local file path" },
      name: { type: "string", description: "Attachment name (default: file name)" },
      contentType: { type: "string", description: "MIME type (default from extension: png/jpg/jpeg/gif/svg/pdf, else application/octet-stream)" },
      isInline: { type: "boolean", description: "Inline attachment (embedded image), default false" },
      contentId: { type: "string", description: "Content-ID referenced from the HTML body as cid:<contentId>" }
    },
    required: ["path"]
  }
};

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent' / 'sentitems', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. HTML emails are securely sanitized to extract only visible text, preventing prompt injection attacks via hidden content.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        includeRawHtml: {
          type: "boolean",
          description: "Include raw HTML content (UNSAFE - for debugging only, may contain hidden prompt injection content)"
        },
        format: {
          type: "string",
          enum: ["text", "html"],
          description: "\"text\" (default): sanitized visible text. \"html\": JSON { id, subject, from, to, cc, sentDateTime, receivedDateTime, bodyContentType, body (unmodified HTML), attachments: [{ id, name, contentType, contentId, isInline, size }] }"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email. Supports both plain text and HTML content.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        },
        attachments: attachmentsSchema
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "draft-email",
    description: "Creates and saves an email draft in Outlook",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Draft email subject"
        },
        body: {
          type: "string",
          description: "Draft email body content (can be plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to save as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        attachments: attachmentsSchema,
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "download-attachments",
    description: "Saves the file attachments of an email to a local folder. Returns JSON: [{ path, name, contentType, contentId, isInline, size }]",
    inputSchema: {
      type: "object",
      properties: {
        messageId: {
          type: "string",
          description: "ID of the email"
        },
        saveDir: {
          type: "string",
          description: "Absolute path of the folder to save into (created if missing; files with the same name are overwritten)"
        },
        inlineOnly: {
          type: "boolean",
          description: "Only save inline attachments (embedded images). Default: false"
        }
      },
      required: ["messageId", "saveDir"]
    },
    handler: handleDownloadAttachments
  }
  // delete-email avregistrerat — verktyget exponeras inte längre via MCP (kan permanent-radera mejl).
  // Implementeringen i email/delete.js och handleDeleteEmail är orörda; kommentera in nedan för att återaktivera.
  /*
  ,{
    name: "delete-email",
    description: "Deletes an email by moving it to Deleted Items (trash). Use permanent=true to hard delete.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to delete"
        },
        permanent: {
          type: "boolean",
          description: "If true, permanently delete the email instead of moving to Deleted Items. Default: false"
        }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  }
  */
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleDraftEmail,
  handleMarkAsRead,
  handleDeleteEmail,
  handleDownloadAttachments
};
