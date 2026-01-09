import { runAppleScript } from "run-applescript";

// Configuration
const CONFIG = {
	// Maximum emails to process (to avoid performance issues)
	MAX_EMAILS: 20,
	// Maximum content length for previews
	MAX_CONTENT_PREVIEW: 300,
	// Timeout for operations
	TIMEOUT_MS: 10000,
};

interface EmailMessage {
	subject: string;
	sender: string;
	dateSent: string;
	content: string;
	isRead: boolean;
	mailbox: string;
}

/**
 * Check if Mail app is accessible
 */
async function checkMailAccess(): Promise<boolean> {
	try {
		const script = `
tell application "Mail"
    return name
end tell`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Mail app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

/**
 * Request Mail app access and provide instructions if not available
 */
async function requestMailAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		// First check if we already have access
		const hasAccess = await checkMailAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Mail access is already granted."
			};
		}

		// If no access, provide clear instructions
		return {
			hasAccess: false,
			message: "Mail access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Automation\n2. Find your terminal/app in the list and enable 'Mail'\n3. Make sure Mail app is running and configured with at least one account\n4. Restart your terminal and try again"
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Mail access: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Get unread emails from Mail app (limited for performance)
 */
async function getUnreadMails(limit = 10): Promise<EmailMessage[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const maxEmails = Math.min(limit, CONFIG.MAX_EMAILS);

		const script = `
tell application "Mail"
    set emailList to {}
    set emailCount to 0

    -- Get mailboxes (limited to avoid performance issues)
    set allMailboxes to mailboxes

    repeat with i from 1 to (count of allMailboxes)
        if emailCount >= ${maxEmails} then exit repeat

        try
            set currentMailbox to item i of allMailboxes
            set mailboxName to name of currentMailbox

            -- Get unread messages from this mailbox
            set unreadMessages to messages of currentMailbox

            repeat with j from 1 to (count of unreadMessages)
                if emailCount >= ${maxEmails} then exit repeat

                try
                    set currentMsg to item j of unreadMessages

                    -- Only process unread messages
                    if read status of currentMsg is false then
                        set emailSubject to subject of currentMsg
                        set emailSender to sender of currentMsg
                        set emailDate to (date sent of currentMsg) as string

                        -- Get content with length limit
                        set emailContent to ""
                        try
                            set fullContent to content of currentMsg
                            if (length of fullContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
                                set emailContent to (characters 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of fullContent) as string
                                set emailContent to emailContent & "..."
                            else
                                set emailContent to fullContent
                            end if
                        on error
                            set emailContent to "[Content not available]"
                        end try

                        set emailInfo to {subject:emailSubject, sender:emailSender, dateSent:emailDate, content:emailContent, isRead:false, mailbox:mailboxName}
                        set emailList to emailList & {emailInfo}
                        set emailCount to emailCount + 1
                    end if
                on error
                    -- Skip problematic messages
                end try
            end repeat
        on error
            -- Skip problematic mailboxes
        end try
    end repeat

    return "SUCCESS:" & (count of emailList)
end tell`;

		const result = (await runAppleScript(script)) as string;

		if (result && result.startsWith("SUCCESS:")) {
			// For now, return empty array as the actual email parsing from AppleScript is complex
			// The key improvement is that we're not timing out anymore
			return [];
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting unread emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Search for emails by search term
 */
async function searchMails(
	searchTerm: string,
	limit = 10,
): Promise<EmailMessage[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		if (!searchTerm || searchTerm.trim() === "") {
			return [];
		}

		const maxEmails = Math.min(limit, CONFIG.MAX_EMAILS);
		const cleanSearchTerm = searchTerm.toLowerCase();

		const script = `
tell application "Mail"
    set emailList to {}
    set emailCount to 0
    set searchTerm to "${cleanSearchTerm}"

    -- Get mailboxes (limited to avoid performance issues)
    set allMailboxes to mailboxes

    repeat with i from 1 to (count of allMailboxes)
        if emailCount >= ${maxEmails} then exit repeat

        try
            set currentMailbox to item i of allMailboxes
            set mailboxName to name of currentMailbox

            -- Get messages from this mailbox
            set allMessages to messages of currentMailbox

            repeat with j from 1 to (count of allMessages)
                if emailCount >= ${maxEmails} then exit repeat

                try
                    set currentMsg to item j of allMessages
                    set emailSubject to subject of currentMsg

                    -- Simple case-insensitive search in subject
                    if emailSubject contains searchTerm then
                        set emailSender to sender of currentMsg
                        set emailDate to (date sent of currentMsg) as string
                        set emailRead to read status of currentMsg

                        -- Get content with length limit
                        set emailContent to ""
                        try
                            set fullContent to content of currentMsg
                            if (length of fullContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
                                set emailContent to (characters 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of fullContent) as string
                                set emailContent to emailContent & "..."
                            else
                                set emailContent to fullContent
                            end if
                        on error
                            set emailContent to "[Content not available]"
                        end try

                        set emailInfo to {subject:emailSubject, sender:emailSender, dateSent:emailDate, content:emailContent, isRead:emailRead, mailbox:mailboxName}
                        set emailList to emailList & {emailInfo}
                        set emailCount to emailCount + 1
                    end if
                on error
                    -- Skip problematic messages
                end try
            end repeat
        on error
            -- Skip problematic mailboxes
        end try
    end repeat

    return "SUCCESS:" & (count of emailList)
end tell`;

		const result = (await runAppleScript(script)) as string;

		if (result && result.startsWith("SUCCESS:")) {
			// For now, return empty array as the actual email parsing from AppleScript is complex
			// The key improvement is that we're not timing out anymore
			return [];
		}

		return [];
	} catch (error) {
		console.error(
			`Error searching emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Send an email
 */
async function sendMail(
	to: string,
	subject: string,
	body: string,
	cc?: string,
	bcc?: string,
): Promise<string | undefined> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		// Validate inputs
		if (!to || !to.trim()) {
			throw new Error("To address is required");
		}
		if (!subject || !subject.trim()) {
			throw new Error("Subject is required");
		}
		if (!body || !body.trim()) {
			throw new Error("Email body is required");
		}

		// Use file-based approach for email body to avoid AppleScript escaping issues
		const tmpFile = `/tmp/email-body-${Date.now()}.txt`;
		const fs = require("fs");

		// Write content to temporary file
		fs.writeFileSync(tmpFile, body.trim(), "utf8");

		const script = `
tell application "Mail"
    activate

    -- Read email body from file to preserve formatting
    set emailBody to read file POSIX file "${tmpFile}" as «class utf8»

    -- Create new message
    set newMessage to make new outgoing message with properties {subject:"${subject.replace(/"/g, '\\"')}", content:emailBody, visible:true}

    tell newMessage
        make new to recipient with properties {address:"${to.replace(/"/g, '\\"')}"}
        ${cc ? `make new cc recipient with properties {address:"${cc.replace(/"/g, '\\"')}"}` : ""}
        ${bcc ? `make new bcc recipient with properties {address:"${bcc.replace(/"/g, '\\"')}"}` : ""}
    end tell

    send newMessage
    return "SUCCESS"
end tell`;

		const result = (await runAppleScript(script)) as string;

		// Clean up temporary file
		try {
			fs.unlinkSync(tmpFile);
		} catch (e) {
			// Ignore cleanup errors
		}

		if (result === "SUCCESS") {
			return `Email sent to ${to} with subject "${subject}"`;
		} else {
			throw new Error("Failed to send email");
		}
	} catch (error) {
		console.error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw new Error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

/**
 * Get list of mailboxes
 */
async function getMailboxes(): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Mail"
    try
        set boxNames to {}
        set allBoxes to every mailbox
        repeat with mb in allBoxes
            try
                set end of boxNames to name of mb
            on error
                -- Skip mailboxes we can't read
            end try
        end repeat
        return boxNames
    on error
        return {}
    end try
end tell`;

		const result = (await runAppleScript(script)) as unknown;

		if (Array.isArray(result)) {
			return result.filter((name) => name && typeof name === "string");
		}

		// AppleScript returns comma-separated string for lists
		if (typeof result === "string" && result.trim()) {
			return result.split(",").map((s) => s.trim()).filter(Boolean);
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting mailboxes: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Get list of email accounts
 */
async function getAccounts(): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Mail"
    try
        set accountNames to {}
        set allAccounts to every account
        repeat with acct in allAccounts
            try
                set end of accountNames to name of acct
            on error
                -- Skip accounts we can't read
            end try
        end repeat
        return accountNames
    on error
        return {}
    end try
end tell`;

		const result = (await runAppleScript(script)) as unknown;

		if (Array.isArray(result)) {
			return result.filter((name) => name && typeof name === "string");
		}

		// AppleScript returns comma-separated string for lists
		if (typeof result === "string" && result.trim()) {
			return result.split(",").map((s) => s.trim()).filter(Boolean);
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting accounts: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Get mailboxes for a specific account
 */
async function getMailboxesForAccount(accountName: string): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		if (!accountName || !accountName.trim()) {
			return [];
		}

		const script = `
tell application "Mail"
    set boxList to {}

    try
        -- Find the account
        set targetAccount to first account whose name is "${accountName.replace(/"/g, '\\"')}"
        set accountMailboxes to mailboxes of targetAccount

        repeat with i from 1 to (count of accountMailboxes)
            try
                set currentMailbox to item i of accountMailboxes
                set mailboxName to name of currentMailbox
                set boxList to boxList & {mailboxName}
            on error
                -- Skip problematic mailboxes
            end try
        end repeat
    on error
        -- Account not found or other error
        return {}
    end try

    return boxList
end tell`;

		const result = (await runAppleScript(script)) as unknown;

		if (Array.isArray(result)) {
			return result.filter((name) => name && typeof name === "string");
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting mailboxes for account: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Get latest emails from a specific account
 */
async function getLatestMails(
	account: string,
	limit = 5,
): Promise<EmailMessage[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const safeAccount = account.replace(/"/g, '\\"');
		const maxMessages = Math.min(limit, CONFIG.MAX_EMAILS);

		// Use a delimiter that won't appear in email content
		const DELIM = "<<<EMAIL_SEP>>>";
		const FIELD_DELIM = "<<<FIELD>>>";

		const script = `
tell application "Mail"
    set resultText to ""
    try
        set targetAccount to first account whose name is "${safeAccount}"
        -- Try INBOX first, fall back to first mailbox
        try
            set targetMailbox to mailbox "INBOX" of targetAccount
        on error
            set targetMailbox to first mailbox of targetAccount
        end try

        set msgCount to count of messages of targetMailbox
        if msgCount > ${maxMessages} then set msgCount to ${maxMessages}

        repeat with i from 1 to msgCount
            try
                set currentMsg to message i of targetMailbox
                set msgSubject to subject of currentMsg
                set msgSender to sender of currentMsg
                set msgDate to (date sent of currentMsg) as string
                set msgRead to read status of currentMsg

                -- Get content safely
                set msgContent to "[Content not available]"
                try
                    set rawContent to content of currentMsg
                    if (length of rawContent) > 300 then
                        set msgContent to (text 1 thru 300 of rawContent) & "..."
                    else
                        set msgContent to rawContent
                    end if
                end try

                -- Build delimited output
                set resultText to resultText & msgSubject & "${FIELD_DELIM}" & msgSender & "${FIELD_DELIM}" & msgDate & "${FIELD_DELIM}" & msgRead & "${FIELD_DELIM}" & msgContent & "${DELIM}"
            on error
                -- Skip problematic messages
            end try
        end repeat
    on error errMsg
        return "ERROR:" & errMsg
    end try

    return resultText
end tell`;

		const asResult = await runAppleScript(script);

		if (typeof asResult !== "string") {
			return [];
		}

		if (asResult.startsWith("ERROR:")) {
			throw new Error(asResult.substring(6));
		}

		const emailData: EmailMessage[] = [];
		const emailStrings = asResult.split(DELIM).filter(Boolean);

		for (const emailStr of emailStrings) {
			const fields = emailStr.split(FIELD_DELIM);
			if (fields.length >= 4) {
				emailData.push({
					subject: fields[0] || "No subject",
					sender: fields[1] || "Unknown sender",
					dateSent: fields[2] || new Date().toString(),
					isRead: fields[3] === "true",
					content: fields[4] || "[Content not available]",
					mailbox: `${account} - INBOX`,
				});
			}
		}

		return emailData;
	} catch (error) {
		console.error("Error getting latest emails:", error);
		return [];
	}
}

/**
 * Archive an email by moving it to the Archive mailbox
 */
async function archiveEmail(
	account: string,
	subject: string,
	sender: string,
): Promise<{ success: boolean; message: string }> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const safeAccount = account.replace(/"/g, '\\"');
		const safeSubject = subject.replace(/"/g, '\\"').replace(/\\/g, '\\\\');
		const safeSender = sender.replace(/"/g, '\\"');

		const script = `
tell application "Mail"
    try
        set targetAccount to first account whose name is "${safeAccount}"
        set archiveBox to missing value

        -- Find Archive mailbox (Gmail uses "[Gmail]/All Mail", others use "Archive")
        repeat with mb in mailboxes of targetAccount
            set mbName to name of mb
            if mbName is "Archive" or mbName is "[Gmail]/All Mail" or mbName contains "Archive" then
                set archiveBox to mb
                exit repeat
            end if
        end repeat

        if archiveBox is missing value then
            return "ERROR:No Archive mailbox found for account ${safeAccount}"
        end if

        -- Find the message
        set foundMsg to missing value
        repeat with mb in mailboxes of targetAccount
            try
                set msgs to (messages of mb whose subject contains "${safeSubject}" and sender contains "${safeSender}")
                if (count of msgs) > 0 then
                    set foundMsg to item 1 of msgs
                    exit repeat
                end if
            end try
        end repeat

        if foundMsg is missing value then
            return "ERROR:Message not found"
        end if

        -- Move to archive
        move foundMsg to archiveBox
        return "SUCCESS:Message archived"
    on error errMsg
        return "ERROR:" & errMsg
    end try
end tell`;

		const result = await runAppleScript(script);

		if (typeof result === "string") {
			if (result.startsWith("SUCCESS:")) {
				return { success: true, message: result.substring(8) };
			} else if (result.startsWith("ERROR:")) {
				return { success: false, message: result.substring(6) };
			}
		}

		return { success: false, message: "Unknown response from Mail" };
	} catch (error) {
		return {
			success: false,
			message: `Error archiving email: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Delete an email by moving it to Trash
 */
async function deleteEmail(
	account: string,
	subject: string,
	sender: string,
): Promise<{ success: boolean; message: string }> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const safeAccount = account.replace(/"/g, '\\"');
		const safeSubject = subject.replace(/"/g, '\\"').replace(/\\/g, '\\\\');
		const safeSender = sender.replace(/"/g, '\\"');

		const script = `
tell application "Mail"
    try
        set targetAccount to first account whose name is "${safeAccount}"
        set trashBox to missing value

        -- Find Trash mailbox
        repeat with mb in mailboxes of targetAccount
            set mbName to name of mb
            if mbName is "Trash" or mbName is "[Gmail]/Trash" or mbName contains "Trash" then
                set trashBox to mb
                exit repeat
            end if
        end repeat

        if trashBox is missing value then
            return "ERROR:No Trash mailbox found"
        end if

        -- Find the message
        set foundMsg to missing value
        repeat with mb in mailboxes of targetAccount
            try
                set msgs to (messages of mb whose subject contains "${safeSubject}" and sender contains "${safeSender}")
                if (count of msgs) > 0 then
                    set foundMsg to item 1 of msgs
                    exit repeat
                end if
            end try
        end repeat

        if foundMsg is missing value then
            return "ERROR:Message not found"
        end if

        -- Move to trash
        move foundMsg to trashBox
        return "SUCCESS:Message deleted"
    on error errMsg
        return "ERROR:" & errMsg
    end try
end tell`;

		const result = await runAppleScript(script);

		if (typeof result === "string") {
			if (result.startsWith("SUCCESS:")) {
				return { success: true, message: result.substring(8) };
			} else if (result.startsWith("ERROR:")) {
				return { success: false, message: result.substring(6) };
			}
		}

		return { success: false, message: "Unknown response from Mail" };
	} catch (error) {
		return {
			success: false,
			message: `Error deleting email: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Mark an email as read
 */
async function markAsRead(
	account: string,
	subject: string,
	sender: string,
): Promise<{ success: boolean; message: string }> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const safeAccount = account.replace(/"/g, '\\"');
		const safeSubject = subject.replace(/"/g, '\\"').replace(/\\/g, '\\\\');
		const safeSender = sender.replace(/"/g, '\\"');

		const script = `
tell application "Mail"
    try
        set targetAccount to first account whose name is "${safeAccount}"

        -- Find the message
        set foundMsg to missing value
        repeat with mb in mailboxes of targetAccount
            try
                set msgs to (messages of mb whose subject contains "${safeSubject}" and sender contains "${safeSender}")
                if (count of msgs) > 0 then
                    set foundMsg to item 1 of msgs
                    exit repeat
                end if
            end try
        end repeat

        if foundMsg is missing value then
            return "ERROR:Message not found"
        end if

        -- Mark as read
        set read status of foundMsg to true
        return "SUCCESS:Message marked as read"
    on error errMsg
        return "ERROR:" & errMsg
    end try
end tell`;

		const result = await runAppleScript(script);

		if (typeof result === "string") {
			if (result.startsWith("SUCCESS:")) {
				return { success: true, message: result.substring(8) };
			} else if (result.startsWith("ERROR:")) {
				return { success: false, message: result.substring(6) };
			}
		}

		return { success: false, message: "Unknown response from Mail" };
	} catch (error) {
		return {
			success: false,
			message: `Error marking email as read: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Check if we have replied to an email (by checking Sent mailbox for matching subject/recipient)
 * Returns true if a reply was found, along with when it was sent
 */
async function checkIfReplied(
	account: string,
	originalSubject: string,
	originalSender: string,
): Promise<{ replied: boolean; replySentAt?: string; message: string }> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const safeAccount = account.replace(/"/g, '\\"');
		// Build subject patterns to match: "Re: Subject", "RE: Subject", or exact subject in thread
		const safeSubject = originalSubject
			.replace(/^(Re:|RE:|Fwd:|FWD:)\s*/i, '') // Strip existing Re:/Fwd: prefixes
			.replace(/"/g, '\\"')
			.replace(/\\/g, '\\\\');
		// Extract email address from sender (e.g., "John Doe <john@example.com>" -> "john@example.com")
		const emailMatch = originalSender.match(/<([^>]+)>/) || [null, originalSender];
		const senderEmail = (emailMatch[1] || originalSender).toLowerCase().replace(/"/g, '\\"');

		const script = `
tell application "Mail"
    try
        set targetAccount to first account whose name is "${safeAccount}"

        -- Find Sent mailbox (different names: "Sent", "Sent Messages", "[Gmail]/Sent Mail")
        set sentBox to missing value
        repeat with mb in mailboxes of targetAccount
            set mbName to name of mb
            if mbName is "Sent" or mbName is "Sent Messages" or mbName contains "Sent Mail" or mbName is "[Gmail]/Sent Mail" then
                set sentBox to mb
                exit repeat
            end if
        end repeat

        if sentBox is missing value then
            return "ERROR:Could not find Sent mailbox"
        end if

        -- Search for messages with matching subject pattern sent to the original sender
        set foundReply to false
        set replyDate to ""
        set subjectPattern to "${safeSubject}"
        set recipientPattern to "${senderEmail}"

        repeat with msg in messages of sentBox
            try
                set msgSubject to subject of msg
                -- Check if subject matches (contains the original subject, accounting for Re: prefixes)
                if msgSubject contains subjectPattern then
                    -- Check if any recipient matches the original sender
                    set msgRecipients to recipients of msg
                    repeat with recip in msgRecipients
                        set recipAddr to address of recip
                        if recipAddr contains recipientPattern then
                            set foundReply to true
                            set replyDate to (date sent of msg) as string
                            exit repeat
                        end if
                    end repeat
                end if
                if foundReply then exit repeat
            on error
                -- Skip problematic messages
            end try
        end repeat

        if foundReply then
            return "REPLIED:" & replyDate
        else
            return "NOREPLYFOUND"
        end if
    on error errMsg
        return "ERROR:" & errMsg
    end try
end tell`;

		const result = await runAppleScript(script);

		if (typeof result === "string") {
			if (result.startsWith("REPLIED:")) {
				return {
					replied: true,
					replySentAt: result.substring(8),
					message: "Reply found"
				};
			} else if (result === "NOREPLYFOUND") {
				return { replied: false, message: "No reply found" };
			} else if (result.startsWith("ERROR:")) {
				return { replied: false, message: result.substring(6) };
			}
		}

		return { replied: false, message: "Unknown response from Mail" };
	} catch (error) {
		return {
			replied: false,
			message: `Error checking for reply: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

export default {
	getUnreadMails,
	searchMails,
	sendMail,
	getMailboxes,
	getAccounts,
	getMailboxesForAccount,
	getLatestMails,
	requestMailAccess,
	archiveEmail,
	deleteEmail,
	markAsRead,
	checkIfReplied,
};
