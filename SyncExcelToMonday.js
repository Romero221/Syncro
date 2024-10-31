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
        width: 1000,
        height: 900,
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

// Modify select-file handler
ipcMain.handle('select-file', async (event, mode) => {
    if (mode === 'open') {
        const result = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
        });
        if (result.canceled) {
            return null;
        } else {
            return result.filePaths[0];
        }
    } else {
        // For 'save' mode (if needed)
        const result = await dialog.showSaveDialog({
            title: 'Save Excel File',
            defaultPath: 'board_data.xlsx',
            filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
        });
        if (result.canceled) {
            return null;
        } else {
            return result.filePath;
        }
    }
});

// Add a listener for syncing Monday to Excel
ipcMain.on('sync-monday-to-excel', async (event, args) => {
    const { boardId, apiKey, filePath } = args;

    try {
        await syncMondayToExcel(boardId, apiKey, filePath);
        event.reply('sync-result', { success: true, message: 'Done!' });
    } catch (error) {
        console.error('An error occurred:', error);
        event.reply('sync-result', { success: false, message: error.message });
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

// Implement the fetchMondayDataForGroup function
async function getGroupId(boardId, apiKey, groupName) {
    const query = `
    query ($boardId: [ID!]!) {
      boards(ids: $boardId) {
        groups {
          id
          title
        }
      }
    }
  `;

    const variables = {
        boardId: [boardId.toString()],
    };

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query, variables },
            {
                headers: { Authorization: `Bearer ${apiKey}` },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', JSON.stringify(response.data.errors, null, 2));
            return null;
        }

        const groups = response.data.data.boards[0].groups;

        // Find the group by title
        const group = groups.find(
            (g) => g.title.trim().toLowerCase() === groupName.trim().toLowerCase()
        );

        if (!group) {
            console.error(`Group "${groupName}" not found on the board.`);
            return null;
        }

        return group.id;
    } catch (error) {
        if (error.response && error.response.data) {
            console.error(
                'Error fetching groups from Monday.com:',
                JSON.stringify(error.response.data, null, 2)
            );
        } else {
            console.error('Error fetching groups from Monday.com:', error.message);
        }
        return null;
    }
}

//################################################################################################################################################

async function fetchItemsByGroup(boardId, groupId, apiKey) {
    let items = [];
    let cursor = null;

    const query = `
    query ($boardId: [ID!]!, $limit: Int, $cursor: String) {
      boards(ids: $boardId) {
        items_page(limit: $limit, cursor: $cursor) {
          cursor
          items {
            id
            name
            group {
              id
            }
            column_values {
              column {
                title
              }
              text
              value
            }
          }
        }
      }
    }
  `;

    const variables = {
        boardId: [boardId.toString()],
        limit: 500,
    };

    try {
        console.log('Starting to fetch items from Monday.com...');
        // Loop until there are no more pages (cursor is null)
        do {
            // Include the cursor in variables if it's not the first page
            if (cursor) {
                variables.cursor = cursor;
                console.log(`Fetching next page with cursor: ${cursor}`);
            } else {
                delete variables.cursor;
                console.log('Fetching first page of items...');
            }

            const response = await axios.post(
                'https://api.monday.com/v2',
                { query, variables },
                {
                    headers: { Authorization: `Bearer ${apiKey}` },
                }
            );

            if (response.data.errors) {
                console.error(
                    'GraphQL errors:',
                    JSON.stringify(response.data.errors, null, 2)
                );
                return null;
            }

            // Append the items from the current page
            const pageItems = response.data.data.boards[0].items_page.items;
            console.log(`Fetched ${pageItems.length} items from current page.`);
            items = items.concat(pageItems);

            // Update the cursor for the next page
            cursor = response.data.data.boards[0].items_page.cursor;

        } while (cursor); // Continue if there's a next page

        console.log(`Total items fetched: ${items.length}`);

        // Filter the items by the specified group ID
        const filteredItems = items.filter(item => item.group.id === groupId);
        console.log(`Items after filtering by group ID "${groupId}": ${filteredItems.length}`);

        if (!filteredItems || filteredItems.length === 0) {
            console.error(`No items found in group "${groupId}".`);
            return null;
        }

        // Log each item's details
        filteredItems.forEach(item => {
            console.log(`Item ID: ${item.id}, Name: ${item.name}`);
            item.column_values.forEach(colVal => {
                const columnTitle = colVal.column ? colVal.column.title : 'Unknown';
                const textValue = colVal.text ? colVal.text : 'No Text';
                console.log(`    Column: ${columnTitle}, Value: ${textValue}`);
            });
        });

        return filteredItems;
    } catch (error) {
        if (error.response && error.response.data) {
            console.error(
                'Error fetching items from Monday.com:',
                JSON.stringify(error.response.data, null, 2)
            );
        } else {
            console.error('Error fetching items from Monday.com:', error.message);
        }
        return null;
    }
}


// Implement the compareAndUpdateExcelData function
function compareAndUpdateExcelData(excelData, mondayData) {
    const updatedExcelData = [...excelData]; // Clone Excel data to avoid direct modifications

    // Create a map of Monday.com items by item name (assuming 'name' is unique within the group)
    const mondayItemsMap = {};
    if (mondayData && mondayData.items) {
        mondayData.items.forEach((item) => {
            mondayItemsMap[item.name.trim()] = item;
        });
    } else {
        console.error("Monday.com data is empty or undefined.");
        return updatedExcelData; // Return unmodified data if Monday data is empty
    }

    // Loop through each row in the Excel data and compare
    updatedExcelData.forEach((excelRow, rowIndex) => {
        const itemName = String(excelRow.Year).trim(); // Assuming 'Year' is a unique identifier
        const mondayItem = mondayItemsMap[itemName];

        if (mondayItem) {
            // Compare each column's value within the row, excluding "Comment" column
            mondayItem.column_values.forEach((colVal) => {
                const columnTitle = colVal.column ? colVal.column.title.trim() : "";
                const mondayValue = colVal.text ? colVal.text.trim() : "";

                // Exclude "Comment" column from the update
                if (columnTitle && columnTitle.toLowerCase() !== "comment") {
                    const excelValue = excelRow[columnTitle] !== null && excelRow[columnTitle] !== undefined
                        ? String(excelRow[columnTitle]).trim()
                        : "";

                    // Update only if values are different
                    if (mondayValue !== excelValue) {
                        console.log(
                            `Updating cell [Row ${rowIndex + 2}][Column: ${columnTitle}]: "${excelValue}" -> "${mondayValue}"`
                        );
                        excelRow[columnTitle] = mondayValue; // Only update value in cloned data
                    }
                }
            });
        } else {
            console.warn(`Item "${itemName}" not found in Monday.com data.`);
        }
    });

    return updatedExcelData; // Return updated data with only changed cells modified
}




// Modify writeDataToExcelFile function if necessary
function writeDataToExcelFile(dataRows, filePath) {
    // Convert data to worksheet
    const worksheet = xlsx.utils.json_to_sheet(dataRows);

    // Create a new workbook and append the worksheet
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Write to file
    xlsx.writeFile(workbook, filePath);
}


// Implement the syncMondayToExcel function
async function syncMondayToExcel(boardId, apiKey, filePath) {
    const groupName = extractManufacturer(filePath); // Extract the group name from the Excel file name
    console.log(`Group Name extracted from file: ${groupName}`);

    // Read the existing Excel file
    console.log('Reading existing Excel file...');
    const excelData = readExcelFile(filePath);

    // Fetch the Group ID
    console.log(`Fetching group ID for group: ${groupName}`);
    const groupId = await getGroupId(boardId, apiKey, groupName);

    if (!groupId) {
        throw new Error('Failed to fetch group ID from Monday.com.');
    }

    console.log(`Group ID for "${groupName}" is ${groupId}`);

    // Fetch items from the group
    console.log(`Fetching items from group ID: ${groupId}`);
    const mondayItems = await fetchItemsByGroup(boardId, groupId, apiKey);

    if (!mondayItems) {
        throw new Error('Failed to fetch items from Monday.com.');
    }

    console.log('Data fetched from Monday.com successfully.');

    // Prepare Monday.com data for comparison
    const mondayData = {
        items: mondayItems,
    };

    // Compare data and update Excel file
    console.log('Comparing data and updating Excel file...');
    const updatedExcelData = compareAndUpdateExcelData(excelData, mondayData);

    // Write updated data back to the Excel file
    writeDataToExcelFile(updatedExcelData, filePath);

    console.log(`Excel file updated successfully at ${filePath}`);
}




//########################################################################################################################


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
