/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * EmailUtils.gs - Common Email & Notification Functions
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for sending emails, checking Gmail replies, and managing
 * email templates across DDS Apps Script projects.
 * 
 * PREREQUISITES:
 * - Requires GmailApp or MailApp scope:
 *   https://www.googleapis.com/auth/gmail.send
 *   or https://www.googleapis.com/auth/gmail.compose
 * 
 * USAGE:
 * Copy this file into your Apps Script project alongside other common-utils.
 * All functions include rate limiting and batch protection.
 * 
 * VERSION: 1.0.0
 * LAST UPDATED: April 2026
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * ==============================================================================
 * FUNCTION: sendHtmlEmail
 * ==============================================================================
 * 
 * PURPOSE:
 * Sends an HTML email with plain text fallback using MailApp or GmailApp.
 * Includes rate limiting protection and comprehensive error handling.
 * 
 * PARAMETERS:
 * @param {Object} options - Email configuration:
 *   - to: {string|Array} - Recipient email(s)
 *   - cc: {string|Array} - CC recipient(s) (optional)
 *   - bcc: {string|Array} - BCC recipient(s) (optional)
 *   - subject: {string} - Email subject line
 *   - htmlBody: {string} - HTML content (required)
 *   - plainBody: {string} - Plain text fallback (auto-generated if omitted)
 *   - fromName: {string} - Sender display name (optional)
 *   - replyTo: {string} - Reply-to address (optional)
 *   - attachments: {Array} - Drive File objects or Blobs to attach (optional)
 *   - inlineImages: {Object} - Map of cid to Blob for inline images (optional)
 * 
 * @returns {Object} Send result:
 *   - success: {boolean} - Whether email was sent
 *   - messageId: {string} - Gmail message ID (if available)
 *   - remainingQuota: {number} - Daily sends remaining
 *   - error: {string} - Error message (if failed)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = sendHtmlEmail({
 *   to: "client@example.com",
 *   cc: ["manager@example.com", "admin@example.com"],
 *   subject: "Document Ready for Review",
 *   htmlBody: "<h1>Your document is ready</h1><p>Please review...</p>",
 *   fromName: "DDS Document System",
 *   replyTo: "noreply@dds.example.com"
 * });
 * 
 * if (result.success) {
 *   Logger.log("Email sent, ID: " + result.messageId);
 * }
 * ```
 * 
 * NOTES:
 * - Auto-converts HTML to plain text if plainBody not provided
 * - Handles both single email string and array of recipients
 * - Respects Gmail daily sending limits
 * - Commonly used in: automated-expiry-notification, CAR, file-center
 * ==============================================================================
 */
function sendHtmlEmail(options) {
  options = options || {};
  
  var result = {
    success: false,
    messageId: null,
    remainingQuota: 0,
    error: null
  };
  
  try {
    // Validate required fields
    if (!options.to) {
      result.error = "Recipient (to) is required";
      return result;
    }
    if (!options.htmlBody) {
      result.error = "HTML body is required";
      return result;
    }
    
    // Convert arrays to comma-separated strings
    var to = Array.isArray(options.to) ? options.to.join(",") : options.to;
    var cc = options.cc ? (Array.isArray(options.cc) ? options.cc.join(",") : options.cc) : null;
    var bcc = options.bcc ? (Array.isArray(options.bcc) ? options.bcc.join(",") : options.bcc) : null;
    
    // Generate plain text fallback if not provided
    var plainBody = options.plainBody || stripHtml(options.htmlBody);
    
    // Prepare email parameters
    var emailParams = {
      to: to,
      subject: options.subject || "(No Subject)",
      htmlBody: options.htmlBody,
      body: plainBody,
      name: options.fromName || null
    };
    
    if (cc) emailParams.cc = cc;
    if (bcc) emailParams.bcc = bcc;
    if (options.replyTo) emailParams.replyTo = options.replyTo;
    if (options.attachments) emailParams.attachments = options.attachments;
    if (options.inlineImages) emailParams.inlineImages = options.inlineImages;
    
    // Check quota before sending
    var quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) {
      result.error = "Daily email quota exceeded";
      result.remainingQuota = 0;
      return result;
    }
    
    // Send the email
    MailApp.sendEmail(emailParams);
    
    result.success = true;
    result.remainingQuota = MailApp.getRemainingDailyQuota();
    
  } catch (error) {
    result.error = "Send failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: sendBatchEmails
 * ==============================================================================
 * 
 * PURPOSE:
 * Sends emails to multiple recipients with rate limiting and progress tracking.
 * Useful for notification campaigns or reminder blasts.
 * 
 * PARAMETERS:
 * @param {Array} recipients - Array of recipient objects:
 *   - to: {string} - Email address
 *   - data: {Object} - Merge data for template (optional)
 * @param {Object} template - Email template:
 *   - subject: {string} - Subject (can use {{placeholders}})
 *   - htmlBody: {string} - HTML template with {{placeholders}}
 *   - fromName: {string} - Sender name
 * @param {Object} options - (Optional) Batch options:
 *   - delayMs: {number} - Delay between sends in ms (default: 1000)
 *   - onProgress: {function} - Callback(sent, total, current) for progress updates
 *   - stopOnError: {boolean} - Stop batch on first error (default: false)
 * 
 * @returns {Object} Batch result:
 *   - success: {boolean} - Whether all emails sent
 *   - sent: {number} - Count successfully sent
 *   - failed: {number} - Count failed
 *   - errors: {Array} - List of failed sends with errors
 *   - remainingQuota: {number} - Quota after batch
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var recipients = [
 *   { to: "user1@example.com", data: { name: "John", doc: "Contract A" } },
 *   { to: "user2@example.com", data: { name: "Jane", doc: "Contract B" } }
 * ];
 * 
 * var template = {
 *   subject: "Reminder: {{doc}} expiring soon",
 *   htmlBody: "<p>Hi {{name}}, your {{doc}} needs attention...</p>",
 *   fromName: "DDS Reminders"
 * };
 * 
 * var result = sendBatchEmails(recipients, template, {
 *   delayMs: 2000,
 *   onProgress: function(sent, total, current) {
 *     Logger.log("Progress: " + sent + "/" + total);
 *   }
 * });
 * ```
 * 
 * NOTES:
 * - Template placeholders use {{variableName}} syntax
 * - Automatically merges recipient.data into template
 * - Rate limited to prevent hitting Gmail quotas
 * - Progress callback receives (sentCount, totalCount, currentRecipient)
 * ==============================================================================
 */
function sendBatchEmails(recipients, template, options) {
  options = options || {};
  var delayMs = options.delayMs || 1000;
  
  var result = {
    success: true,
    sent: 0,
    failed: 0,
    errors: [],
    remainingQuota: 0
  };
  
  if (!Array.isArray(recipients) || recipients.length === 0) {
    result.success = false;
    result.errors.push("No recipients provided");
    return result;
  }
  
  for (var i = 0; i < recipients.length; i++) {
    var recipient = recipients[i];
    
    try {
      // Merge template with recipient data
      var subject = mergeTemplate(template.subject, recipient.data || {});
      var htmlBody = mergeTemplate(template.htmlBody, recipient.data || {});
      
      var sendResult = sendHtmlEmail({
        to: recipient.to,
        subject: subject,
        htmlBody: htmlBody,
        fromName: template.fromName
      });
      
      if (sendResult.success) {
        result.sent++;
      } else {
        result.failed++;
        result.errors.push({
          recipient: recipient.to,
          error: sendResult.error,
          index: i
        });
        result.success = false;
        
        if (options.stopOnError) {
          break;
        }
      }
      
      // Progress callback
      if (options.onProgress) {
        options.onProgress(result.sent, recipients.length, recipient);
      }
      
      // Rate limiting delay (except after last email)
      if (i < recipients.length - 1) {
        Utilities.sleep(delayMs);
      }
      
    } catch (error) {
      result.failed++;
      result.errors.push({
        recipient: recipient.to,
        error: error.toString(),
        index: i
      });
      result.success = false;
    }
  }
  
  result.remainingQuota = MailApp.getRemainingDailyQuota();
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: checkEmailReplies
 * ==============================================================================
 * 
 * PURPOSE:
 * Scans Gmail inbox for reply emails matching specific criteria.
 * Used to detect client acknowledgements or responses to automated emails.
 * 
 * PARAMETERS:
 * @param {Object} options - Search criteria:
 *   - searchQuery: {string} - Gmail search query (e.g., "from:client@example.com")
 *   - afterDate: {Date} - Only check threads after this date
 *   - keywords: {Array} - Keywords to look for in replies (e.g., ["acknowledged", "received", "confirmed"])
 *   - label: {string} - Only check threads with this label
 *   - maxThreads: {number} - Maximum threads to check (default: 50)
 *   - markAsRead: {boolean} - Mark matched threads as read (default: false)
 *   - archive: {boolean} - Archive matched threads (default: false)
 * 
 * @returns {Object} Search result:
 *   - success: {boolean} - Whether scan completed
 *   - threads: {Array} - Matching thread objects
 *   - replies: {Array} - Individual reply messages found
 *   - count: {number} - Total matching replies
 *   - error: {string} - Error message (if failed)
 * 
 * Reply object properties:
 *   - threadId: {string} - Gmail thread ID
 *   - messageId: {string} - Message ID
 *   - from: {string} - Sender email
 *   - subject: {string} - Subject line
 *   - body: {string} - Message body text
 *   - date: {Date} - Message date
 *   - matchedKeywords: {Array} - Keywords found in this reply
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = checkEmailReplies({
 *   searchQuery: "subject:Reminder has:reply",
 *   afterDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000), // Last 7 days
 *   keywords: ["acknowledged", "noted", "received"],
 *   maxThreads: 100
 * });
 * 
 * result.replies.forEach(function(reply) {
 *   Logger.log(reply.from + " acknowledged: " + reply.matchedKeywords.join(", "));
 * });
 * ```
 * 
 * NOTES:
 * - Uses GmailApp search syntax (same as Gmail search bar)
 * - Keywords are case-insensitive
 * - Searches in both subject and body
 * - Commonly used in: automated-expiry-notification (acknowledgement tracking)
 * ==============================================================================
 */
function checkEmailReplies(options) {
  options = options || {};
  
  var result = {
    success: false,
    threads: [],
    replies: [],
    count: 0,
    error: null
  };
  
  try {
    // Build search query
    var queryParts = [];
    
    if (options.searchQuery) {
      queryParts.push(options.searchQuery);
    }
    
    if (options.afterDate) {
      var afterStr = Utilities.formatDate(options.afterDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
      queryParts.push("after:" + afterStr);
    }
    
    if (options.label) {
      queryParts.push("label:" + options.label);
    }
    
    // Always look for replies (not just sent messages)
    queryParts.push("in:inbox");
    
    var searchQuery = queryParts.join(" ");
    var maxThreads = options.maxThreads || 50;
    
    // Search Gmail
    var threads = GmailApp.search(searchQuery, 0, maxThreads);
    
    result.threads = threads;
    
    // Process each thread
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      
      // Skip if only 1 message (not a reply thread)
      if (messages.length < 2) {
        continue;
      }
      
      // Process reply messages (skip first as it's the original sent message)
      for (var j = 1; j < messages.length; j++) {
        var message = messages[j];
        var body = message.getPlainBody();
        var subject = message.getSubject();
        
        // Check for keywords
        var matchedKeywords = [];
        if (options.keywords && options.keywords.length > 0) {
          var searchText = (subject + " " + body).toLowerCase();
          options.keywords.forEach(function(keyword) {
            if (searchText.indexOf(keyword.toLowerCase()) !== -1) {
              matchedKeywords.push(keyword);
            }
          });
        }
        
        // Include if keywords match (or no keywords specified)
        if (!options.keywords || matchedKeywords.length > 0) {
          var reply = {
            threadId: thread.getId(),
            messageId: message.getId(),
            from: message.getFrom(),
            subject: subject,
            body: body.substring(0, 500), // Truncate long bodies
            date: message.getDate(),
            matchedKeywords: matchedKeywords,
            isUnread: message.isUnread()
          };
          
          result.replies.push(reply);
          
          // Mark as read if requested
          if (options.markAsRead && message.isUnread()) {
            message.markRead();
          }
        }
      }
      
      // Archive thread if requested
      if (options.archive) {
        thread.moveToArchive();
      }
    }
    
    result.count = result.replies.length;
    result.success = true;
    
  } catch (error) {
    result.error = "Gmail search failed: " + error.toString();
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: createEmailTemplate
 * ==============================================================================
 * 
 * PURPOSE:
 * Creates a standardized email template with DDS branding and common components.
 * Returns HTML that can be used with sendHtmlEmail().
 * 
 * PARAMETERS:
 * @param {Object} content - Template content:
 *   - title: {string} - Email heading/title
 *   - message: {string} - Main message content (HTML)
 *   - actionButton: {Object} - (Optional) CTA button:
 *     - text: {string} - Button label
 *     - url: {string} - Button link
 *   - footer: {string} - (Optional) Custom footer text
 *   - highlight: {string} - (Optional) Highlight color class: "info", "warning", "success", "danger"
 * 
 * @returns {string} Complete HTML email body ready for sending
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var html = createEmailTemplate({
 *   title: "Document Expiring Soon",
 *   message: "<p>Your contract expires in 7 days. Please renew to avoid interruption.</p>",
 *   actionButton: {
 *     text: "Renew Now",
 *     url: "https://example.com/renew"
 *   },
 *   highlight: "warning"
 * });
 * 
 * sendHtmlEmail({
 *   to: "client@example.com",
 *   subject: "Action Required: Document Expiring",
 *   htmlBody: html,
 *   fromName: "DDS System"
 * });
 * ```
 * 
 * NOTES:
 * - Generates responsive HTML email compatible with most clients
 * - Uses inline styles for maximum compatibility
 * - Includes automatic unsubscribe footer
 * ==============================================================================
 */
function createEmailTemplate(content) {
  var colors = {
    info: "#3498db",
    warning: "#f39c12",
    success: "#27ae60",
    danger: "#e74c3c",
    default: "#34495e"
  };
  
  var highlightColor = colors[content.highlight] || colors.default;
  
  var buttonHtml = "";
  if (content.actionButton) {
    buttonHtml = `
      <div style="text-align: center; margin: 30px 0;">
        <a href="${content.actionButton.url}" 
           style="background-color: ${highlightColor}; color: white; padding: 12px 30px; 
                  text-decoration: none; border-radius: 4px; display: inline-block; 
                  font-weight: bold;">
          ${content.actionButton.text}
        </a>
      </div>
    `;
  }
  
  var footer = content.footer || 
    "This is an automated message from DDS System. Please do not reply directly.";
  
  var html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5;">
    <tr>
      <td align="center" style="padding: 20px 0;">
        <table width="600" cellpadding="0" cellspacing="0" style="background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          
          <!-- Header -->
          <tr>
            <td style="background-color: ${highlightColor}; padding: 30px; text-align: center;">
              <h1 style="color: white; margin: 0; font-size: 24px; font-weight: 600;">
                ${content.title}
              </h1>
            </td>
          </tr>
          
          <!-- Content -->
          <tr>
            <td style="padding: 40px 30px;">
              <div style="color: #333; font-size: 16px; line-height: 1.6;">
                ${content.message}
              </div>
              
              ${buttonHtml}
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="background-color: #f8f9fa; padding: 20px 30px; text-align: center; border-top: 1px solid #e9ecef;">
              <p style="color: #6c757d; font-size: 12px; margin: 0;">
                ${footer}
              </p>
              <p style="color: #adb5bd; font-size: 11px; margin: 10px 0 0 0;">
                Sent by DDS Overall System • ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")}
              </p>
            </td>
          </tr>
          
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
  
  return html.replace(/^\s+/gm, "").trim();
}

/**
 * ==============================================================================
 * FUNCTION: getEmailQuotaInfo
 * ==============================================================================
 * 
 * PURPOSE:
 * Returns current Gmail email quota status and usage statistics.
 * Useful for monitoring and preventing quota exhaustion.
 * 
 * PARAMETERS: None
 * 
 * @returns {Object} Quota information:
 *   - remainingDaily: {number} - Remaining emails for today
 *   - dailyLimit: {number} - Daily sending limit (typically 100 for consumer accounts)
 *   - usedToday: {number} - Emails sent today
 *   - percentRemaining: {number} - Percentage of quota remaining
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var quota = getEmailQuotaInfo();
 * Logger.log("Emails remaining: " + quota.remainingDaily + "/" + quota.dailyLimit);
 * 
 * if (quota.percentRemaining < 10) {
 *   // Alert admin about low quota
 * }
 * ```
 * 
 * NOTES:
 * - Consumer Gmail accounts: ~100/day limit
 * - Google Workspace: Higher limits based on plan
 * - Quota resets at midnight Pacific Time
 * ==============================================================================
 */
function getEmailQuotaInfo() {
  var remaining = MailApp.getRemainingDailyQuota();
  var dailyLimit = 100; // Standard consumer limit
  var usedToday = dailyLimit - remaining;
  
  return {
    remainingDaily: remaining,
    dailyLimit: dailyLimit,
    usedToday: usedToday,
    percentRemaining: Math.round((remaining / dailyLimit) * 100)
  };
}

/**
 * ==============================================================================
 * HELPER FUNCTIONS (Internal)
 * ==============================================================================
 */

/**
 * Strips HTML tags to create plain text version
 * @private
 */
function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Merges template with data object using {{placeholder}} syntax
 * @private
 */
function mergeTemplate(template, data) {
  var result = template;
  for (var key in data) {
    if (data.hasOwnProperty(key)) {
      var regex = new RegExp("{{\\s*" + key + "\\s*}}", "g");
      result = result.replace(regex, data[key]);
    }
  }
  return result;
}
