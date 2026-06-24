import { Task, TaskUpdate, User, TaskStatus, UserRole } from '../types';

interface SpreadsheetResponse {
  spreadsheetId: string;
  spreadsheetUrl: string;
}

/**
 * Creates a new blank Google Spreadsheet with formatted tabs.
 */
export const createExportSpreadsheet = async (
  accessToken: string,
  title: string
): Promise<SpreadsheetResponse> => {
  const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      properties: {
        title: title,
      },
      sheets: [
        { properties: { sheetId: 0, title: 'Overview' } },
        { properties: { sheetId: 1, title: 'Tasks Checklist' } },
        { properties: { sheetId: 2, title: 'Updates Timeline' } },
        { properties: { sheetId: 3, title: 'Team Directory' } },
      ],
    }),
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Failed to create spreadsheet: ${errText}`);
  }

  const data = await response.json();
  return {
    spreadsheetId: data.spreadsheetId,
    spreadsheetUrl: data.spreadsheetUrl,
  };
};

/**
 * Populates all tabs with the application's task, history, and team data in a single batch.
 */
export const populateSpreadsheetData = async (
  accessToken: string,
  spreadsheetId: string,
  data: {
    tasks: Task[];
    updates: TaskUpdate[];
    users: User[];
    currentUser: User;
  }
): Promise<void> => {
  const { tasks, updates, users, currentUser } = data;

  // Build high-efficiency username mapper
  const userMap = new Map<string, string>();
  users.forEach(u => userMap.set(u.id, u.username));

  // Build high-efficiency task title mapper
  const taskMap = new Map<string, string>();
  tasks.forEach(t => taskMap.set(t.id, t.title));

  const timestamp = new Date().toLocaleString();

  // 1. Overview Sheet Data
  const overviewValues = [
    ['GROW DAILY - WORKSPACE DATA EXPORT', ''],
    ['----------------------------------', ''],
    ['Export Timestamp', timestamp],
    ['Exported By', currentUser.username],
    ['User Role', currentUser.role],
    ['', ''],
    ['WORKSPACE METRICS STATS', ''],
    ['----------------------------------', ''],
    ['Total Registered Tasks', tasks.length],
    ['Total Status Updates Logs', updates.length],
    ['Total Workspace Users', users.length],
  ];

  // 2. Tasks Checklist Sheet Data
  const tasksHeaders = [
    'Task ID',
    'Assignee / Owner',
    'Task Title',
    'Category',
    'Current Status',
    'Target Due Date',
    'Created At Time',
    'Associated Tags',
  ];
  const tasksValues = [
    tasksHeaders,
    ...tasks.map(t => [
      t.id,
      userMap.get(t.userId) || t.userId,
      t.title,
      t.category,
      t.status,
      t.targetDate,
      new Date(t.createdAt).toLocaleString(),
      t.tags ? t.tags.join(', ') : '',
    ]),
  ];

  // 3. Updates Timeline Sheet Data
  const updatesHeaders = [
    'Update ID',
    'Related Task ID',
    'Related Task Title',
    'Date Logged',
    'Content / Note',
    'Status Transition',
    'Logged Timestamp',
  ];
  const updatesValues = [
    updatesHeaders,
    ...updates.map(u => [
      u.id,
      u.taskId,
      taskMap.get(u.taskId) || 'N/A (Deleted Task)',
      u.date,
      u.content,
      u.statusChange ? `${u.statusChange.from} ➔ ${u.statusChange.to}` : 'None',
      new Date(u.timestamp).toLocaleString(),
    ]),
  ];

  // 4. Team Directory Sheet Data
  const teamHeaders = [
    'User ID',
    'Username',
    'Role Type',
    'Managed Team Line (Member IDs)',
  ];
  const teamValues = [
    teamHeaders,
    ...users.map(u => [
      u.id,
      u.username,
      u.role,
      u.teamMemberIds ? u.teamMemberIds.join(', ') : '',
    ]),
  ];

  // Send batch update request for cell values
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: "'Overview'!A1", values: overviewValues },
          { range: "'Tasks Checklist'!A1", values: tasksValues },
          { range: "'Updates Timeline'!A1", values: updatesValues },
          { range: "'Team Directory'!A1", values: teamValues },
        ],
      }),
    }
  );

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Failed to populate spreadsheet data: ${errText}`);
  }

  // Optional: Apply formatting/styles to headers (making row 1 font bold, setting alternating background row styles etc.)
  // We can call spreadsheets BatchUpdate to format sheets nicely
  const formatResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        requests: [
          // Basic font family & bold formatting for all 4 sheets' header row
          ...[0, 1, 2, 3].map(sheetIndex => ({
            repeatCell: {
              range: {
                sheetId: sheetIndex, // Note: the sheetId in created spreadsheet is generated sequentially if created this way
                startRowIndex: 0,
                endRowIndex: 1,
              },
              cell: {
                userEnteredFormat: {
                  textFormat: {
                    bold: true,
                    fontSize: 11,
                  },
                },
              },
              fields: 'userEnteredFormat.textFormat',
            },
          })),
        ],
      }),
    }
  ).catch(err => {
    // Suppress minor formatting errors if Google generate sequential sheetIds slightly differently
    console.warn('Formatting spreadsheet headers encountered a secondary issue:', err);
  });
};

interface ImportedData {
  tasks: Task[];
  updates: TaskUpdate[];
  users: User[];
}

/**
 * Parses and extracts data from a Google Spreadsheet by calling batchGet.
 */
export const importSpreadsheetData = async (
  accessToken: string,
  spreadsheetId: string,
  currentUserId?: string
): Promise<ImportedData> => {
  const ranges = encodeURIComponent('Tasks Checklist!A1:Z5000') + 
                 '&ranges=' + encodeURIComponent('Updates Timeline!A1:Z5000') + 
                 '&ranges=' + encodeURIComponent('Team Directory!A1:Z5000');
                 
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchGet?ranges=${ranges}`,
    {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Failed to read spreadsheet data: ${errText}`);
  }

  const result = await response.json();
  const valueRanges = result.valueRanges || [];

  const tasksRange = valueRanges.find((vr: any) => vr.range && vr.range.includes('Tasks Checklist'));
  const updatesRange = valueRanges.find((vr: any) => vr.range && vr.range.includes('Updates Timeline'));
  const teamRange = valueRanges.find((vr: any) => vr.range && vr.range.includes('Team Directory'));

  const parsedUsers: User[] = [];
  const parsedTasks: Task[] = [];
  const parsedUpdates: TaskUpdate[] = [];

  // Parse Team/Users first so we can resolve assignee usernames in tasks
  const userMap = new Map<string, string>(); // Username -> User ID
  
  if (teamRange && teamRange.values && teamRange.values.length > 1) {
    const rows = teamRange.values;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0] || !row[1]) continue;
      
      const userId = row[0].trim();
      const username = row[1].trim();
      const role = row[2]?.trim() === 'ADMIN' ? 'ADMIN' : 'MEMBER';
      const teamMemberIds = row[3] ? row[3].split(',').map((s: string) => s.trim()).filter(Boolean) : [];
      
      parsedUsers.push({
        id: userId,
        username,
        pin: '1234', // Default placeholder PIN, existing PINs are preserved in database merge
        role,
        teamMemberIds,
      });
      
      userMap.set(username.toLowerCase(), userId);
    }
  }

  // Parse Tasks
  if (tasksRange && tasksRange.values && tasksRange.values.length > 1) {
    const rows = tasksRange.values;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0] || !row[2]) continue; // ID and Title are required

      const id = row[0].trim();
      const assigneeInput = row[1] ? row[1].trim() : '';
      const title = row[2].trim();
      const category = row[3] ? row[3].trim() : 'General';
      
      let status = TaskStatus.NOT_YET;
      const statusInput = row[4] ? row[4].trim().toUpperCase() : '';
      if (Object.values(TaskStatus).includes(statusInput as TaskStatus)) {
        status = statusInput as TaskStatus;
      }

      const targetDate = row[5] ? row[5].trim() : new Date().toISOString().slice(0, 10);
      
      let createdAt = Date.now();
      if (row[6]) {
        const parsedTime = Date.parse(row[6].trim());
        if (!isNaN(parsedTime)) {
          createdAt = parsedTime;
        }
      }

      const tags = row[7] ? row[7].split(',').map((s: string) => s.trim()).filter(Boolean) : [];

      // Determine userId
      let userId = '';
      if (assigneeInput) {
        const testAssignee = assigneeInput.toLowerCase();
        if (userMap.has(testAssignee)) {
          userId = userMap.get(testAssignee)!;
        } else {
          // Generate a new user profile for this missing assignee
          const newId = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            const r = Math.random() * 16 | 0;
            const v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
          });
          parsedUsers.push({
            id: newId,
            username: assigneeInput,
            pin: '1234',
            role: 'MEMBER',
            teamMemberIds: []
          });
          userMap.set(testAssignee, newId);
          userId = newId;
        }
      } else if (currentUserId) {
        userId = currentUserId;
      }

      parsedTasks.push({
        id,
        userId,
        title,
        category,
        status,
        targetDate,
        createdAt,
        tags,
      });
    }
  }

  // Parse Updates Tracker
  if (updatesRange && updatesRange.values && updatesRange.values.length > 1) {
    const rows = updatesRange.values;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0] || !row[1]) continue; // ID and taskId are required

      const id = row[0].trim();
      const taskId = row[1].trim();
      const date = row[3] ? row[3].trim() : new Date().toISOString().slice(0, 10);
      const content = row[4] ? row[4].trim() : '';

      let statusChange;
      const transitionInput = row[5] ? row[5].trim() : '';
      if (transitionInput && transitionInput !== 'None') {
        const parts = transitionInput.split(/➔|->/).map((s: string) => s.trim().toUpperCase());
        if (parts.length === 2) {
          const from = parts[0] as TaskStatus;
          const to = parts[1] as TaskStatus;
          if (Object.values(TaskStatus).includes(from) && Object.values(TaskStatus).includes(to)) {
            statusChange = { from, to };
          }
        }
      }

      let timestamp = Date.now();
      if (row[6]) {
        const parsedTime = Date.parse(row[6].trim());
        if (!isNaN(parsedTime)) {
          timestamp = parsedTime;
        }
      }

      const updateDoc: TaskUpdate = {
        id,
        taskId,
        date,
        content,
        timestamp,
        isArchived: false,
      };

      if (statusChange) {
        updateDoc.statusChange = statusChange;
      }

      parsedUpdates.push(updateDoc);
    }
  }

  return {
    tasks: parsedTasks,
    updates: parsedUpdates,
    users: parsedUsers,
  };
};

/**
 * Helper to parse a standard CSV file string into arrays.
 */
export const parseCSV = (text: string): string[][] => {
  const result = [];
  let row: string[] = [];
  let col = '';
  let inQuotes = false;
  
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        col += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      row.push(col);
      col = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      row.push(col);
      col = '';
      if (row.length > 0 && (row.length > 1 || row[0] !== '')) {
        result.push(row);
      }
      row = [];
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
    } else {
      col += char;
    }
  }
  if (col !== '' || row.length > 0) {
    row.push(col);
    result.push(row);
  }
  return result;
};

