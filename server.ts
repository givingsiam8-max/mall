import express, { Request, Response } from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import * as XLSX from 'xlsx';
import { getDb, queryAll, queryOne, runSql, saveDbImmediate } from './server/db';
import { parseSpreadsheetBuffer } from './server/excelParser';
import {
  loadCurrentSettings,
  testEmailProviderConnection,
  verifyUnsubscribeToken,
  EmailSettings,
} from './server/emailService';
import { CampaignQueueManager } from './server/campaignQueue';

const app = express();
const PORT = Number(process.env.PORT || 3000);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB file size limit
});

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// Helper to get Base URL
function getBaseUrl(req: Request): string {
  if (process.env.APP_URL) return process.env.APP_URL;
  const protocol = req.headers['x-forwarded-proto'] || req.protocol;
  const host = req.headers['x-forwarded-host'] || req.get('host');
  return `${protocol}://${host}`;
}

// -------------------------------------------------------------
// Public Unsubscribe Routes
// -------------------------------------------------------------
app.get('/unsubscribe', (req: Request, res: Response) => {
  const email = (req.query.email as string) || '';
  const campaignId = Number(req.query.c) || 0;
  const token = (req.query.token as string) || '';

  // Render a clean, standalone, responsive unsubscribe page
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Unsubscribe Confirmation</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            margin: 0;
            padding: 20px;
          }
          .card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            max-width: 480px;
            width: 100%;
            padding: 36px 32px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            text-align: center;
          }
          .icon {
            width: 48px;
            height: 48px;
            background: #f1f5f9;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 20px;
            color: #475569;
          }
          h1 {
            font-size: 20px;
            font-weight: 600;
            margin: 0 0 12px 0;
            color: #0f172a;
          }
          p {
            font-size: 14px;
            line-height: 1.5;
            color: #64748b;
            margin: 0 0 24px 0;
          }
          .email-tag {
            font-family: monospace;
            background: #f1f5f9;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 13px;
            color: #334155;
            display: inline-block;
            margin-bottom: 20px;
          }
          button {
            background: #0f172a;
            color: #ffffff;
            border: none;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: 500;
            border-radius: 6px;
            cursor: pointer;
            width: 100%;
          }
          button:hover {
            background: #1e293b;
          }
          .success-msg {
            color: #16a34a;
            font-weight: 500;
            display: none;
            margin-top: 16px;
          }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M18 6 6 18"/><path d="m6 6 12 12"/>
            </svg>
          </div>
          <h1>Manage Your Subscription</h1>
          <p>Confirm that you wish to unsubscribe from all marketing and campaign emails sent to:</p>
          <div class="email-tag">${email || 'your address'}</div>
          <form id="unsub-form">
            <input type="hidden" name="email" value="${email}" />
            <input type="hidden" name="campaignId" value="${campaignId}" />
            <input type="hidden" name="token" value="${token}" />
            <button type="submit" id="submit-btn">Unsubscribe Me</button>
          </form>
          <div id="result" class="success-msg">You have been unsubscribed successfully.</div>
        </div>
        <script>
          const form = document.getElementById('unsub-form');
          const btn = document.getElementById('submit-btn');
          const result = document.getElementById('result');

          form.addEventListener('submit', async (e) => {
            e.preventDefault();
            btn.disabled = true;
            btn.textContent = 'Processing...';

            try {
              const res = await fetch('/api/unsubscribe', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  email: '${email}',
                  campaignId: ${campaignId},
                  token: '${token}'
                })
              });
              const data = await res.json();
              if (data.success) {
                btn.style.display = 'none';
                result.style.display = 'block';
                result.textContent = 'You have been unsubscribed successfully. Your address is suppressed from all future campaigns.';
              } else {
                alert(data.message || 'Failed to unsubscribe');
                btn.disabled = false;
                btn.textContent = 'Unsubscribe Me';
              }
            } catch (err) {
              alert('Error connecting to server');
              btn.disabled = false;
              btn.textContent = 'Unsubscribe Me';
            }
          });
        </script>
      </body>
    </html>
  `);
});

app.post('/api/unsubscribe', (req: Request, res: Response) => {
  const { email, campaignId, token } = req.body;
  if (!email || typeof email !== 'string') {
    return res.status(400).json({ success: false, message: 'Valid email required' });
  }

  const cleanEmail = email.trim().toLowerCase();
  const now = new Date().toISOString();

  // If token is provided, verify it (or allow direct unsubscribe)
  if (campaignId && token) {
    const isValid = verifyUnsubscribeToken(cleanEmail, Number(campaignId), token);
    if (!isValid) {
      // Still allow unsubscribe but log token mismatch
      console.warn(`Unsubscribe token mismatch for email ${cleanEmail}`);
    }
  }

  // 1. Add to suppression_list
  runSql(
    'INSERT OR IGNORE INTO suppression_list (email, reason, created_at) VALUES (?, ?, ?)',
    [cleanEmail, 'unsubscribe', now]
  );

  // 2. Update contacts table
  runSql(
    "UPDATE contacts SET status = 'unsubscribed', unsubscribe_status = 'unsubscribed', suppression_status = 'unsubscribe' WHERE email = ?",
    [cleanEmail]
  );

  // 3. Update any campaign recipient status
  if (campaignId) {
    runSql(
      "UPDATE campaign_recipients SET status = 'unsubscribed' WHERE campaign_id = ? AND email = ?",
      [Number(campaignId), cleanEmail]
    );
  }

  return res.json({ success: true, message: 'Unsubscribed successfully' });
});

// -------------------------------------------------------------
// Dashboard Statistics API
// -------------------------------------------------------------
app.get('/api/stats', (req: Request, res: Response) => {
  const totalContacts = queryOne<{ count: number }>('SELECT COUNT(*) as count FROM contacts')?.count || 0;
  const activeContacts = queryOne<{ count: number }>("SELECT COUNT(*) as count FROM contacts WHERE status = 'active'")?.count || 0;
  const suppressedContacts = queryOne<{ count: number }>("SELECT COUNT(*) as count FROM contacts WHERE status = 'suppressed'")?.count || 0;
  const unsubscribedContacts = queryOne<{ count: number }>("SELECT COUNT(*) as count FROM contacts WHERE status = 'unsubscribed'")?.count || 0;
  
  const totalCampaigns = queryOne<{ count: number }>('SELECT COUNT(*) as count FROM campaigns')?.count || 0;
  const totalSent = queryOne<{ count: number }>("SELECT COUNT(*) as count FROM campaign_recipients WHERE status = 'sent'")?.count || 0;
  const totalFailed = queryOne<{ count: number }>("SELECT COUNT(*) as count FROM campaign_recipients WHERE status = 'failed'")?.count || 0;

  const recentImports = queryAll('SELECT * FROM imports ORDER BY id DESC LIMIT 5');
  const recentCampaigns = queryAll('SELECT * FROM campaigns ORDER BY id DESC LIMIT 5');

  // Date breakdown for chart/calendar
  const dateBreakdown = queryAll<{ import_date: string; total: number; active: number }>(`
    SELECT 
      import_date, 
      COUNT(*) as total, 
      SUM(CASE WHEN status = 'active' THEN 1 ELSE 0 END) as active 
    FROM contacts 
    GROUP BY import_date 
    ORDER BY import_date DESC 
    LIMIT 10
  `);

  return res.json({
    totalContacts,
    activeContacts,
    suppressedContacts,
    unsubscribedContacts,
    totalCampaigns,
    totalSent,
    totalFailed,
    recentImports,
    recentCampaigns,
    dateBreakdown,
  });
});

// -------------------------------------------------------------
// Excel / Spreadsheet Upload & Deterministic Extraction
// -------------------------------------------------------------
app.post('/api/import/upload', upload.single('file'), (req: Request, res: Response) => {
  if (!req.file) {
    return res.status(400).json({ success: false, message: 'No file uploaded' });
  }

  const filename = req.file.originalname || 'unknown.xlsx';
  const buffer = req.file.buffer;

  try {
    const parseResult = parseSpreadsheetBuffer(buffer, filename);

    const now = new Date();
    const importedAt = now.toISOString();
    const importDate = now.toISOString().split('T')[0];
    const importTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false });

    // Insert new contacts into SQLite database
    for (const c of parseResult.extractedEmails) {
      runSql(
        `INSERT OR IGNORE INTO contacts 
        (email, source_file, imported_at, import_date, import_time, status, unsubscribe_status, suppression_status, created_at) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          c.email,
          c.sourceFile,
          importedAt,
          importDate,
          importTime,
          c.status,
          c.unsubscribeStatus,
          c.suppressionStatus,
          importedAt,
        ]
      );
    }

    // Save import audit entry
    const insertImport = runSql(
      `INSERT INTO imports 
      (filename, imported_at, import_date, import_time, total_rows, cells_scanned, emails_found, new_emails, duplicates, invalid_candidates, status) 
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        filename,
        importedAt,
        importDate,
        importTime,
        parseResult.totalRows,
        parseResult.cellsScanned,
        parseResult.emailsFound,
        parseResult.newEmails,
        parseResult.duplicateCount,
        parseResult.invalidCandidates,
        'completed',
      ]
    );

    saveDbImmediate();

    return res.json({
      success: true,
      importId: insertImport.lastInsertRowid,
      filename,
      importDate,
      importTime,
      totalRows: parseResult.totalRows,
      cellsScanned: parseResult.cellsScanned,
      emailsFound: parseResult.emailsFound,
      newEmails: parseResult.newEmails,
      duplicateCount: parseResult.duplicateCount,
      inFileDuplicates: parseResult.inFileDuplicates,
      inDbDuplicates: parseResult.inDbDuplicates,
      invalidCandidates: parseResult.invalidCandidates,
    });
  } catch (err: any) {
    console.error('Import processing error:', err);
    return res.status(500).json({
      success: false,
      message: `Failed to process spreadsheet: ${err?.message || 'Invalid or corrupt file'}`,
    });
  }
});

// -------------------------------------------------------------
// Import History API
// -------------------------------------------------------------
app.get('/api/import/history', (req: Request, res: Response) => {
  const imports = queryAll('SELECT * FROM imports ORDER BY id DESC');
  return res.json({ imports });
});

app.delete('/api/import/:id', (req: Request, res: Response) => {
  const id = Number(req.params.id);
  const deleteContacts = req.query.deleteContacts === 'true';

  const imp = queryOne<{ filename: string }>('SELECT filename FROM imports WHERE id = ?', [id]);
  if (!imp) {
    return res.status(404).json({ success: false, message: 'Import record not found' });
  }

  if (deleteContacts) {
    runSql('DELETE FROM contacts WHERE source_file = ?', [imp.filename]);
  }

  runSql('DELETE FROM imports WHERE id = ?', [id]);
  return res.json({ success: true, message: 'Import deleted successfully' });
});

// -------------------------------------------------------------
// Contact Inventory API (Paginated, Searchable, Filterable)
// -------------------------------------------------------------
app.get('/api/contacts', (req: Request, res: Response) => {
  const page = Math.max(1, Number(req.query.page || 1));
  const limit = Math.min(500, Math.max(10, Number(req.query.limit || 25)));
  const offset = (page - 1) * limit;

  const search = ((req.query.search as string) || '').trim();
  const status = (req.query.status as string) || 'all';
  const sourceFile = (req.query.source_file as string) || '';
  const dateFilter = (req.query.date as string) || '';
  const datePreset = (req.query.date_preset as string) || '';
  const sortBy = (req.query.sort_by as string) || 'id';
  const sortOrder = (req.query.sort_order as string) === 'asc' ? 'ASC' : 'DESC';

  const whereConditions: string[] = [];
  const params: any[] = [];

  if (search) {
    whereConditions.push('(email LIKE ? OR source_file LIKE ?)');
    params.push(`%${search}%`, `%${search}%`);
  }

  if (status && status !== 'all') {
    whereConditions.push('status = ?');
    params.push(status);
  }

  if (sourceFile) {
    whereConditions.push('source_file = ?');
    params.push(sourceFile);
  }

  if (dateFilter) {
    whereConditions.push('import_date = ?');
    params.push(dateFilter);
  } else if (datePreset && datePreset !== 'all') {
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];

    if (datePreset === 'today') {
      whereConditions.push('import_date = ?');
      params.push(todayStr);
    } else if (datePreset === 'yesterday') {
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);
      const yesterdayStr = yesterday.toISOString().split('T')[0];
      whereConditions.push('import_date = ?');
      params.push(yesterdayStr);
    } else if (datePreset === 'last_7_days') {
      const past = new Date(today);
      past.setDate(past.getDate() - 7);
      whereConditions.push('import_date >= ?');
      params.push(past.toISOString().split('T')[0]);
    } else if (datePreset === 'last_30_days') {
      const past = new Date(today);
      past.setDate(past.getDate() - 30);
      whereConditions.push('import_date >= ?');
      params.push(past.toISOString().split('T')[0]);
    }
  }

  const whereSql = whereConditions.length > 0 ? `WHERE ${whereConditions.join(' AND ')}` : '';

  // Allowed sort columns
  const allowedSortCols: Record<string, string> = {
    id: 'id',
    email: 'email',
    source_file: 'source_file',
    imported_at: 'imported_at',
    import_date: 'import_date',
    status: 'status',
    last_sent_at: 'last_sent_at',
  };
  const safeSortCol = allowedSortCols[sortBy] || 'id';

  // Count total matching
  const countRow = queryOne<{ count: number }>(`SELECT COUNT(*) as count FROM contacts ${whereSql}`, params);
  const total = countRow?.count || 0;

  // Query records
  const contacts = queryAll(
    `SELECT * FROM contacts ${whereSql} ORDER BY ${safeSortCol} ${sortOrder} LIMIT ? OFFSET ?`,
    [...params, limit, offset]
  );

  return res.json({
    contacts,
    total,
    page,
    limit,
    totalPages: Math.ceil(total / limit),
  });
});

// Grouping by Date
app.get('/api/contacts/dates', (req: Request, res: Response) => {
  const dates = queryAll<{ import_date: string; total: number; active: number }>(`
    SELECT 
      import_date,
      COUNT(*) as total,
      SUM(CASE WHEN status = 'active' THEN 1 ELSE 0 END) as active
    FROM contacts
    GROUP BY import_date
    ORDER BY import_date DESC
  `);
  return res.json({ dates });
});

// Grouping by Source Files
app.get('/api/contacts/sources', (req: Request, res: Response) => {
  const sources = queryAll<{ source_file: string; total: number }>(`
    SELECT source_file, COUNT(*) as total
    FROM contacts
    GROUP BY source_file
    ORDER BY total DESC
  `);
  return res.json({ sources });
});

// Bulk and Single Actions on Contacts
app.post('/api/contacts/suppress', (req: Request, res: Response) => {
  const { ids, emails, reason } = req.body;
  const suppressionReason = reason || 'manual';
  const now = new Date().toISOString();

  let targetEmails: string[] = [];
  if (Array.isArray(emails) && emails.length > 0) {
    targetEmails = emails.map((e) => String(e).trim().toLowerCase());
  } else if (Array.isArray(ids) && ids.length > 0) {
    const placeholders = ids.map(() => '?').join(',');
    const rows = queryAll<{ email: string }>(`SELECT email FROM contacts WHERE id IN (${placeholders})`, ids);
    targetEmails = rows.map((r) => r.email.toLowerCase());
  }

  for (const email of targetEmails) {
    runSql('INSERT OR IGNORE INTO suppression_list (email, reason, created_at) VALUES (?, ?, ?)', [
      email,
      suppressionReason,
      now,
    ]);
    runSql("UPDATE contacts SET status = 'suppressed', suppression_status = ? WHERE email = ?", [
      suppressionReason,
      email,
    ]);
  }

  return res.json({ success: true, count: targetEmails.length });
});

app.post('/api/contacts/unsubscribe', (req: Request, res: Response) => {
  const { ids, emails } = req.body;
  const now = new Date().toISOString();

  let targetEmails: string[] = [];
  if (Array.isArray(emails) && emails.length > 0) {
    targetEmails = emails.map((e) => String(e).trim().toLowerCase());
  } else if (Array.isArray(ids) && ids.length > 0) {
    const placeholders = ids.map(() => '?').join(',');
    const rows = queryAll<{ email: string }>(`SELECT email FROM contacts WHERE id IN (${placeholders})`, ids);
    targetEmails = rows.map((r) => r.email.toLowerCase());
  }

  for (const email of targetEmails) {
    runSql('INSERT OR IGNORE INTO suppression_list (email, reason, created_at) VALUES (?, ?, ?)', [
      email,
      'unsubscribe',
      now,
    ]);
    runSql(
      "UPDATE contacts SET status = 'unsubscribed', unsubscribe_status = 'unsubscribed', suppression_status = 'unsubscribe' WHERE email = ?",
      [email]
    );
  }

  return res.json({ success: true, count: targetEmails.length });
});

app.delete('/api/contacts', (req: Request, res: Response) => {
  const { ids, all } = req.body;
  if (all === true) {
    runSql('DELETE FROM contacts');
    return res.json({ success: true, message: 'All contacts cleared' });
  }

  if (Array.isArray(ids) && ids.length > 0) {
    const placeholders = ids.map(() => '?').join(',');
    runSql(`DELETE FROM contacts WHERE id IN (${placeholders})`, ids);
    return res.json({ success: true, count: ids.length });
  }

  return res.status(400).json({ success: false, message: 'No target IDs provided' });
});

// -------------------------------------------------------------
// Contact Export API (XLSX, CSV, TXT)
// -------------------------------------------------------------
app.post('/api/contacts/export', (req: Request, res: Response) => {
  const { format = 'csv', ids, filter } = req.body;

  let contacts: any[] = [];
  if (Array.isArray(ids) && ids.length > 0) {
    const placeholders = ids.map(() => '?').join(',');
    contacts = queryAll(`SELECT * FROM contacts WHERE id IN (${placeholders}) ORDER BY id ASC`, ids);
  } else {
    // Export based on filters or entire inventory
    contacts = queryAll('SELECT * FROM contacts ORDER BY id ASC');
  }

  const timestamp = new Date().toISOString().split('T')[0];

  if (format === 'txt') {
    // One email per line cleanly
    const txtContent = contacts.map((c) => c.email).join('\r\n') + '\r\n';
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="email-inventory-${timestamp}.txt"`);
    return res.send(txtContent);
  }

  if (format === 'csv') {
    const rows = [
      ['ID', 'Email', 'Source File', 'Imported Date', 'Imported Time', 'Status', 'Last Sent Date'],
      ...contacts.map((c) => [
        c.id,
        c.email,
        c.source_file,
        c.import_date,
        c.import_time,
        c.status,
        c.last_sent_at || '',
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const csvContent = XLSX.utils.sheet_to_csv(ws);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="email-inventory-${timestamp}.csv"`);
    return res.send(csvContent);
  }

  // Default XLSX
  const rows = [
    ['ID', 'Email', 'Source File', 'Imported Date', 'Imported Time', 'Status', 'Last Sent Date'],
    ...contacts.map((c) => [
      c.id,
      c.email,
      c.source_file,
      c.import_date,
      c.import_time,
      c.status,
      c.last_sent_at || '',
    ]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Email Inventory');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="email-inventory-${timestamp}.xlsx"`);
  return res.send(buffer);
});

// -------------------------------------------------------------
// Campaign Management API
// -------------------------------------------------------------
app.get('/api/campaigns', (req: Request, res: Response) => {
  const campaigns = queryAll('SELECT * FROM campaigns ORDER BY id DESC');
  return res.json({ campaigns });
});

app.post('/api/campaigns', (req: Request, res: Response) => {
  const {
    name,
    subject,
    senderName,
    senderEmail,
    replyTo,
    htmlBody,
    textBody,
    recipientSelection,
  } = req.body;

  if (!name || !subject || !htmlBody) {
    return res.status(400).json({ success: false, message: 'Name, Subject, and Email Body are required' });
  }

  const settings = loadCurrentSettings();
  const effectiveSenderName = senderName || settings.fromName || 'Shadow game';
  const effectiveSenderEmail = senderEmail || settings.fromEmail || 'givingsiam8@gmail.com';
  const effectiveReplyTo = replyTo || settings.replyTo || 'givingsiam8@gmail.com';
  const now = new Date().toISOString();

  // Recipient resolution logic
  let eligibleContacts: Array<{ id: number; email: string }> = [];

  const selectionType = recipientSelection?.type || 'all_active';

  // Subquery to exclude suppressed emails
  const suppressionCondition = `email NOT IN (SELECT email FROM suppression_list)`;

  if (selectionType === 'by_ids' && Array.isArray(recipientSelection.ids) && recipientSelection.ids.length > 0) {
    const placeholders = recipientSelection.ids.map(() => '?').join(',');
    eligibleContacts = queryAll<{ id: number; email: string }>(
      `SELECT id, email FROM contacts WHERE id IN (${placeholders}) AND status = 'active' AND ${suppressionCondition}`,
      recipientSelection.ids
    );
  } else if (selectionType === 'by_date' && recipientSelection.date) {
    eligibleContacts = queryAll<{ id: number; email: string }>(
      `SELECT id, email FROM contacts WHERE import_date = ? AND status = 'active' AND ${suppressionCondition}`,
      [recipientSelection.date]
    );
  } else if (selectionType === 'by_source' && recipientSelection.sourceFile) {
    eligibleContacts = queryAll<{ id: number; email: string }>(
      `SELECT id, email FROM contacts WHERE source_file = ? AND status = 'active' AND ${suppressionCondition}`,
      [recipientSelection.sourceFile]
    );
  } else {
    // all active contacts
    eligibleContacts = queryAll<{ id: number; email: string }>(
      `SELECT id, email FROM contacts WHERE status = 'active' AND ${suppressionCondition}`
    );
  }

  // Create Campaign Record
  const totalCount = eligibleContacts.length;
  const insertCamp = runSql(
    `INSERT INTO campaigns 
    (name, subject, sender_name, sender_email, reply_to, html_body, text_body, created_at, status, total_recipients, pending_count, sent_count, failed_count, suppressed_count) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0)`,
    [
      name,
      subject,
      effectiveSenderName,
      effectiveSenderEmail,
      effectiveReplyTo,
      htmlBody,
      textBody || '',
      now,
      'draft',
      totalCount,
      totalCount,
    ]
  );

  const campaignId = insertCamp.lastInsertRowid;

  // Insert recipients
  for (const c of eligibleContacts) {
    runSql(
      `INSERT INTO campaign_recipients (campaign_id, contact_id, email, status) VALUES (?, ?, ?, 'pending')`,
      [campaignId, c.id, c.email]
    );
  }

  saveDbImmediate();

  return res.json({
    success: true,
    campaignId,
    totalRecipients: totalCount,
    message: `Campaign created with ${totalCount} verified recipients`,
  });
});

app.get('/api/campaigns/:id', (req: Request, res: Response) => {
  const id = Number(req.params.id);
  const campaign = queryOne('SELECT * FROM campaigns WHERE id = ?', [id]);
  if (!campaign) {
    return res.status(404).json({ success: false, message: 'Campaign not found' });
  }

  // Breakdown of recipient statuses
  const counts = queryAll<{ status: string; count: number }>(
    'SELECT status, COUNT(*) as count FROM campaign_recipients WHERE campaign_id = ? GROUP BY status',
    [id]
  );

  // Paginated recipient log
  const page = Math.max(1, Number(req.query.page || 1));
  const limit = 20;
  const offset = (page - 1) * limit;

  const recipients = queryAll(
    'SELECT * FROM campaign_recipients WHERE campaign_id = ? ORDER BY id ASC LIMIT ? OFFSET ?',
    [id, limit, offset]
  );

  const totalRecipientsCount = queryOne<{ count: number }>(
    'SELECT COUNT(*) as count FROM campaign_recipients WHERE campaign_id = ?',
    [id]
  )?.count || 0;

  const jobStatus = CampaignQueueManager.getJobStatus(id);

  return res.json({
    campaign,
    recipientCounts: counts,
    recipients,
    pagination: {
      page,
      limit,
      total: totalRecipientsCount,
      totalPages: Math.ceil(totalRecipientsCount / limit),
    },
    jobStatus,
  });
});

// Campaign Controls
app.post('/api/campaigns/:id/start', async (req: Request, res: Response) => {
  const id = Number(req.params.id);
  CampaignQueueManager.setAppUrl(getBaseUrl(req));
  const result = await CampaignQueueManager.startCampaign(id);
  return res.json(result);
});

app.post('/api/campaigns/:id/pause', async (req: Request, res: Response) => {
  const id = Number(req.params.id);
  const result = await CampaignQueueManager.pauseCampaign(id);
  return res.json(result);
});

app.post('/api/campaigns/:id/resume', async (req: Request, res: Response) => {
  const id = Number(req.params.id);
  CampaignQueueManager.setAppUrl(getBaseUrl(req));
  const result = await CampaignQueueManager.resumeCampaign(id);
  return res.json(result);
});

app.post('/api/campaigns/:id/cancel', async (req: Request, res: Response) => {
  const id = Number(req.params.id);
  const result = await CampaignQueueManager.cancelCampaign(id);
  return res.json(result);
});

// -------------------------------------------------------------
// Suppression List API
// -------------------------------------------------------------
app.get('/api/suppression', (req: Request, res: Response) => {
  const list = queryAll('SELECT * FROM suppression_list ORDER BY id DESC');
  return res.json({ suppressionList: list });
});

app.post('/api/suppression', (req: Request, res: Response) => {
  const { email, reason } = req.body;
  if (!email || !email.includes('@')) {
    return res.status(400).json({ success: false, message: 'Valid email required' });
  }

  const clean = email.trim().toLowerCase();
  const now = new Date().toISOString();
  const suppressionReason = reason || 'manual';

  runSql('INSERT OR IGNORE INTO suppression_list (email, reason, created_at) VALUES (?, ?, ?)', [
    clean,
    suppressionReason,
    now,
  ]);
  runSql("UPDATE contacts SET status = 'suppressed', suppression_status = ? WHERE email = ?", [
    suppressionReason,
    clean,
  ]);

  return res.json({ success: true, message: 'Address suppressed' });
});

app.delete('/api/suppression/:id', (req: Request, res: Response) => {
  const id = Number(req.params.id);
  const entry = queryOne<{ email: string }>('SELECT email FROM suppression_list WHERE id = ?', [id]);
  if (entry) {
    runSql('DELETE FROM suppression_list WHERE id = ?', [id]);
    runSql("UPDATE contacts SET status = 'active', suppression_status = 'none' WHERE email = ?", [entry.email]);
  }
  return res.json({ success: true, message: 'Removed from suppression list' });
});

// -------------------------------------------------------------
// Settings & Provider Configuration API
// -------------------------------------------------------------
app.get('/api/settings', (req: Request, res: Response) => {
  const settings = loadCurrentSettings();

  // Mask sensitive values
  const safeSettings = {
    ...settings,
    smtpPassword: settings.smtpPassword ? '••••••••' : '',
    apiKey: settings.apiKey ? `••••${settings.apiKey.slice(-4)}` : '',
  };

  return res.json({ settings: safeSettings });
});

app.post('/api/settings', (req: Request, res: Response) => {
  const current = loadCurrentSettings();
  const incoming = req.body;

  // Preserve existing password or key if masked or empty
  let newPassword = incoming.smtpPassword;
  if (!newPassword || newPassword === '••••••••') {
    newPassword = current.smtpPassword;
  }

  let newApiKey = incoming.apiKey;
  if (!newApiKey || newApiKey.startsWith('••••')) {
    newApiKey = current.apiKey;
  }

  const updated: EmailSettings = {
    provider: incoming.provider || current.provider,
    smtpHost: incoming.smtpHost ?? current.smtpHost,
    smtpPort: Number(incoming.smtpPort || current.smtpPort || 587),
    smtpUser: incoming.smtpUser ?? current.smtpUser,
    smtpPassword: newPassword,
    smtpSecure: incoming.smtpSecure ?? current.smtpSecure,
    apiKey: newApiKey,
    mailgunDomain: incoming.mailgunDomain ?? current.mailgunDomain,
    fromName: incoming.fromName ?? current.fromName,
    fromEmail: incoming.fromEmail ?? current.fromEmail,
    replyTo: incoming.replyTo ?? current.replyTo,
    batchSize: Number(incoming.batchSize || current.batchSize || 50),
    batchDelayMs: Number(incoming.batchDelayMs || current.batchDelayMs || 1000),
    maxRetries: Number(incoming.maxRetries || current.maxRetries || 3),
    dailyLimit: Number(incoming.dailyLimit || current.dailyLimit || 10000),
  };

  runSql('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)', [
    'email_config',
    JSON.stringify(updated),
  ]);
  saveDbImmediate();

  return res.json({ success: true, message: 'Settings saved successfully' });
});

app.post('/api/settings/test-connection', async (req: Request, res: Response) => {
  const { targetEmail } = req.body;
  if (!targetEmail || !targetEmail.includes('@')) {
    return res.status(400).json({ success: false, message: 'Valid test recipient email is required' });
  }

  const currentSettings = loadCurrentSettings();
  const appUrl = getBaseUrl(req);

  const result = await testEmailProviderConnection(targetEmail, currentSettings, appUrl);
  return res.json(result);
});

// -------------------------------------------------------------
// Server Start & Vite Middleware Integration
// -------------------------------------------------------------
async function startServer() {
  // Ensure DB initialized
  await getDb();
  console.log('Database initialized successfully.');

  const isProduction = process.env.NODE_ENV === 'production';

  if (!isProduction) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(process.cwd(), 'dist');
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
      app.get('*', (req: Request, res: Response) => {
        res.sendFile(path.join(distPath, 'index.html'));
      });
    }
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server listening on port ${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Fatal error starting server:', err);
});
