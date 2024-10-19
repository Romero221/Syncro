// SyncExcelToMonday.js

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const fs = require('fs');

let mainWindow;

// Custom logging function
function setupLogging() {
    const log = console.log;
    console.log = function (...args) {
        log.apply(console, args); // Log to the console as usual

        // Send log messages to renderer
        if (mainWindow && mainWindow.webContents) {
            mainWindow.webContents.send('log-message', args.join(' '));
        }
    };
}

// Function to create the main window
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 600,
        height: 500,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            // nodeIntegration: true, // Not needed with preload script
            // contextIsolation: false, // We use contextBridge in preload.js
        },
    });

    mainWindow.loadFile('index.html');

    // Open DevTools for debugging (optional)
    // mainWindow.webContents.openDevTools();

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

// When the app is ready, create the window
app.whenReady().then(() => {
    createWindow();
    setupLogging(); // Initialize custom logging

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

// Quit when all windows are closed
app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

// IPC handlers for communication between renderer and main processes
ipcMain.handle('select-file', async () => {
    const result = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    });
    if (result.canceled) {
        return null;
    } else {
        return result.filePaths[0];
    }
});

ipcMain.on('start-processing', async (event, args) => {
    const { boardId, apiKey, filePath } = args;

    try {
        await updateOrCreateBoard(filePath, apiKey, boardId);
        event.reply('processing-result', { success: true, message: 'Done!' });
    } catch (error) {
        console.error('An error occurred:', error);
        event.reply('processing-result', { success: false, message: error.message });
    }
});

// Your existing functions start here

// Function to read Excel data
function readExcelFile(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
}

// Function to extract the manufacturer (first word) from the file name
function extractManufacturer(fileName) {
    const baseName = path.basename(fileName); // Get the base file name
    return baseName.split(' ')[0]; // Assuming the manufacturer is the first word
}

// Function to fetch existing groups on the board
async function fetchGroups(boardId, apiKey) {
    const query = `
    query($boardId: ID!) {
      boards(ids: [$boardId]) {
        groups {
          id
          title
        }
      }
    }
  `;

    const variables = { boardId: parseInt(boardId) };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return [];
        }

        return response.data.data.boards[0].groups;
    } catch (error) {
        console.error('Error fetching groups:', error);
        return [];
    }
}

// Function to create a new group on the board
async function createGroup(boardId, groupTitle, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $groupTitle: String!) {
      create_group(board_id: $boardId, group_name: $groupTitle) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        groupTitle,
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        return response.data.data.create_group.id;
    } catch (error) {
        console.error('Error creating group:', error);
    }
}

// Function to delete a group from the board
async function deleteGroup(boardId, groupId, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $groupId: String!) {
      delete_group(board_id: $boardId, group_id: $groupId) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        groupId: groupId,
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        console.log(`Deleted group with ID: ${groupId}`);
        return response.data.data.delete_group.id;
    } catch (error) {
        console.error('Error deleting group:', error);
        return null;
    }
}

// Function to fetch existing columns on the board
async function fetchColumns(boardId, apiKey) {
    const query = `
    query($boardId: ID!) {
      boards(ids: [$boardId]) {
        columns {
          id
          title
        }
      }
    }
  `;

    const variables = { boardId: parseInt(boardId) };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return [];
        }

        return response.data.data.boards[0].columns;
    } catch (error) {
        console.error('Error fetching columns:', error);
        return [];
    }
}

// Function to create a new column on the board
async function createColumn(boardId, columnTitle, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $columnTitle: String!, $columnType: ColumnType!) {
      create_column(board_id: $boardId, title: $columnTitle, column_type: $columnType) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        columnTitle,
        columnType: 'text',
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        return response.data.data.create_column.id;
    } catch (error) {
        console.error('Error creating column:', error);
    }
}

// Function to delete a column from the board
async function deleteColumn(boardId, columnId, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $columnId: String!) {
      delete_column(board_id: $boardId, column_id: $columnId) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        columnId: columnId,
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        console.log(`Deleted column with ID: ${columnId}`);
        return response.data.data.delete_column.id;
    } catch (error) {
        console.error('Error deleting column:', error);
    }
}

// Function to create a new item in a specific group on the board (with year as item name)
async function createItemInGroup(boardId, groupId, itemName, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $groupId: String!, $itemName: String!) {
      create_item(board_id: $boardId, group_id: $groupId, item_name: $itemName) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        groupId: groupId,
        itemName: itemName,
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        if (
            response.data &&
            response.data.data &&
            response.data.data.create_item
        ) {
            return response.data.data.create_item.id;
        } else {
            console.error('Unexpected response structure:', response.data);
            return null;
        }
    } catch (error) {
        console.error('Error creating new item in group:', error);
        throw new Error('Failed to create item in group');
    }
}

// Function to update the created item with column values (filling the rest of the row)
async function updateItem(boardId, itemId, columnValues, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $itemId: ID!, $columnValues: JSON!) {
      change_multiple_column_values(board_id: $boardId, item_id: $itemId, column_values: $columnValues) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        itemId: parseInt(itemId),
        columnValues: JSON.stringify(columnValues),
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        if (
            response.data &&
            response.data.data &&
            response.data.data.change_multiple_column_values
        ) {
            return response.data.data.change_multiple_column_values.id;
        } else {
            console.error('Unexpected response structure:', response.data);
            return null;
        }
    } catch (error) {
        console.error('Error updating item:', error);
        throw new Error('Failed to update item');
    }
}

// Function to archive a group on the board
async function archiveGroup(boardId, groupId, apiKey) {
    const mutation = `
    mutation($boardId: ID!, $groupId: String!) {
      archive_group(board_id: $boardId, group_id: $groupId) {
        id
      }
    }
  `;

    const variables = {
        boardId: parseInt(boardId),
        groupId: groupId,
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return null;
        }

        console.log(`Archived group with ID: ${groupId}`);
        return response.data.data.archive_group.id;
    } catch (error) {
        console.error('Error archiving group:', error);
        return null;
    }
}


async function updateOrCreateBoard(filePath, apiKey, boardId) {
    const excelData = readExcelFile(filePath);
    const groupName = extractManufacturer(filePath); // Use the first word of the file name as the group title

    // Fetch existing groups on the board
    const groups = await fetchGroups(boardId, apiKey);

    // Check if the group already exists
    let groupId;
    const existingGroup = groups.find(
        (group) =>
            group.title.trim().toLowerCase() === groupName.trim().toLowerCase()
    );
    if (existingGroup) {
        // Archive the existing group
        console.log(
            `Group "${groupName}" already exists with ID: ${existingGroup.id}. Archiving group.`
        );
        await archiveGroup(boardId, existingGroup.id, apiKey);
    }

    // Create the group
    groupId = await createGroup(boardId, groupName, apiKey);
    if (!groupId) {
        console.error('Failed to create group.');
        return;
    }
    console.log(`Created new group: ${groupName} with ID: ${groupId}`);

    // Fetch existing columns on the board
    let existingColumns = await fetchColumns(boardId, apiKey);
    let existingColumnTitles = existingColumns.map((col) =>
        col.title.trim().toLowerCase()
    );

    console.log('Existing columns:', existingColumnTitles);

    // No need to create the "Year" column as it's the default item name column
    const columnIdMap = {}; // To store the mapping between column titles and their IDs in Monday.com
    console.log('Skipping "Year" column creation as it is the item name column.');

    // Create or match the remaining columns from the Excel sheet (starting from the 2nd column onward)
    const excelColumns = Object.keys(excelData[0]).map((col) => col.trim()); // Get column headers from Excel file and trim them

    for (let i = 1; i < excelColumns.length; i++) {
        // Skip the first "Year" column, as it was handled separately
        const excelColumn = excelColumns[i];
        if (!existingColumnTitles.includes(excelColumn.toLowerCase())) {
            // Create column if it doesn't exist
            const newColumnId = await createColumn(boardId, excelColumn, apiKey);
            columnIdMap[excelColumn] = newColumnId;
            console.log(`Created column: ${excelColumn} with ID: ${newColumnId}`);
        } else {
            // Map existing column if it already exists
            const existingColumn = existingColumns.find(
                (col) =>
                    col.title.trim().toLowerCase() === excelColumn.toLowerCase()
            );
            columnIdMap[excelColumn] = existingColumn.id;
            console.log(
                `Column "${excelColumn}" already exists with ID: ${existingColumn.id}`
            );
        }
    }

    // Ensure 'Comment' column exists at the very end and is not deleted
    const commentColumnTitle = 'Comment';
    let commentColumn = existingColumns.find(
        (col) =>
            col.title.trim().toLowerCase() === commentColumnTitle.toLowerCase()
    );
    if (!commentColumn) {
        // Create 'Comment' column if it doesn't exist
        const newCommentColumnId = await createColumn(
            boardId,
            commentColumnTitle,
            apiKey
        );
        console.log(`Created 'Comment' column with ID: ${newCommentColumnId}`);
        commentColumn = { id: newCommentColumnId, title: commentColumnTitle };
        existingColumns.push(commentColumn);
        existingColumnTitles.push(commentColumnTitle.toLowerCase());
    } else {
        console.log(`'Comment' column already exists with ID: ${commentColumn.id}`);
    }

    // Update existingColumns and existingColumnTitles to include the new 'Comment' column
    existingColumns = await fetchColumns(boardId, apiKey);
    existingColumnTitles = existingColumns.map((col) =>
        col.title.trim().toLowerCase()
    );

    // Delete columns that exist on the board but are not in the Excel sheet, excluding 'Name' and 'Comment' columns
    for (let existingColumn of existingColumns) {
        if (
            existingColumn.title.trim().toLowerCase() !== 'name' && // Do not delete the 'Name' column
            existingColumn.title.trim().toLowerCase() !== 'comment' && // Do not delete the 'Comment' column
            !excelColumns.some(
                (excelCol) =>
                    excelCol.toLowerCase() === existingColumn.title.trim().toLowerCase()
            )
        ) {
            console.log(`Deleting extra column: ${existingColumn.title}`);
            await deleteColumn(boardId, existingColumn.id, apiKey);
        }
    }

    // Iterate over Excel data and create items using the year value as the item name
    for (const row of excelData) {
        const year = String(row.Year).trim(); // Ensure 'Year' is used as the item name (default column)
        const columnValues = {};

        // Map Excel data to Monday.com columns using the correct column IDs, skipping the 'Year' column
        for (let i = 1; i < excelColumns.length; i++) {
            const column = excelColumns[i].trim();
            columnValues[columnIdMap[column]] = String(row[column]).trim();
        }

        // Create new item
        console.log(`Creating new item for year ${year}`);
        const newItemId = await createItemInGroup(
            boardId,
            groupId,
            year,
            apiKey
        );
        console.log(`Created new item for year: ${year} (ID: ${newItemId})`);

        // Wait 1 second before updating the item with the rest of the column values
        await new Promise((resolve) => setTimeout(resolve, 1000));

        // Update the item with the rest of the column values
        console.log(`Updating item ID ${newItemId} with column values.`);
        await updateItem(boardId, newItemId, columnValues, apiKey);
        console.log(`Updated item ID ${newItemId} with column values.`);
    }
}


