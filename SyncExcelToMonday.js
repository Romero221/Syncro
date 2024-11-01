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

    console.log(`Updating item ID: ${itemId} with values:`, columnValues);

    try {
        const response = await axios.post(
            'https://api.monday.com/v2',
            { query: mutation, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors during update:', response.data.errors);
            return null;
        }

        return response.data.data.change_multiple_column_values.id;
    } catch (error) {
        console.error('Error updating item in Monday:', error.message);
        throw new Error('Failed to update item in Monday');
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
function compareAndUpdateExcelData(headers, existingData, newData) {
    const excelItemsMap = {};
    existingData.forEach((row) => {
        const year = String(row['Year']).trim();
        if (year) {
            excelItemsMap[year] = row;
        }
    });

    newData.forEach((mondayRow) => {
        const year = String(mondayRow['Year']).trim();
        if (excelItemsMap[year]) {
            headers.forEach((header) => {
                if (header.toLowerCase() === 'comment') {
                    return; // Skip 'Comment' column
                }

                const excelValue = excelItemsMap[year][header] !== undefined && excelItemsMap[year][header] !== null
                    ? String(excelItemsMap[year][header]).trim()
                    : '';

                const mondayValue = mondayRow[header] !== undefined && mondayRow[header] !== null
                    ? String(mondayRow[header]).trim()
                    : '';

                if (mondayValue !== excelValue) {
                    console.log(`Updating cell [Year: ${year}][Column: ${header}]: "${excelValue}" -> "${mondayValue}"`);
                    excelItemsMap[year][header] = mondayValue;
                }
            });
        } else {
            console.warn(`Item "${year}" not found in Monday.com data.`);
        }
    });

    return existingData; // Return updated data with only changed cells modified
}




// Modify writeDataToExcelFile function if necessary
function writeDataToExcelFile(workbook, sheetName, sheet, headers, dataRows, filePath) {
    // Determine the starting row (assuming headers are in the first row)
    const startRow = 2;

    // Map headers to column indices
    const headerMap = {};
    headers.forEach((header, index) => {
        headerMap[header] = index;
    });

    // Find the maximum row count to preserve existing rows
    const existingRowCount = xlsx.utils.decode_range(sheet['!ref']).e.r;

    // Update existing rows
    dataRows.forEach((rowData, rowIndex) => {
        const excelRowIndex = startRow + rowIndex;
        headers.forEach((header, colIndex) => {
            // Exclude 'Comment' column
            if (header.toLowerCase() === 'comment') {
                return;
            }

            // Get the cell address
            const cellAddress = xlsx.utils.encode_cell({ r: excelRowIndex - 1, c: colIndex });
            let cell = sheet[cellAddress];

            if (cell) {
                const newValue = rowData[header] !== undefined && rowData[header] !== null
                    ? String(rowData[header]).trim()
                    : '';

                // Only update if the value is different
                if (cell.v !== newValue) {
                    console.log(`Updating cell [Row ${excelRowIndex}][Column: ${header}]: "${cell.v}" -> "${newValue}"`);
                    cell.v = newValue;
                    // If you want to ensure the cell type is string
                    cell.t = 's';
                }
            } else {
                // If the cell doesn't exist, create it
                const newValue = rowData[header] !== undefined && rowData[header] !== null
                    ? String(rowData[header]).trim()
                    : '';
                console.log(`Creating cell [Row ${excelRowIndex}][Column: ${header}]: "${newValue}"`);
                sheet[cellAddress] = { t: 's', v: newValue };
            }
        });
    });

    // Adjust the sheet's range if new rows have been added
    const newRowCount = dataRows.length + startRow - 1;
    if (newRowCount > existingRowCount) {
        const newRef = xlsx.utils.encode_range({
            s: { r: 0, c: 0 },
            e: { r: newRowCount - 1, c: headers.length - 1 },
        });
        sheet['!ref'] = newRef;
    }

    // Write the updated workbook back to the file
    xlsx.writeFile(workbook, filePath);
}

//#########################################################

function readExcelFileWithFormatting(filePath) {
    const workbook = xlsx.readFile(filePath, { cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Get the headers (column titles) from the first row
    const headersRow = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    const headers = headersRow[0].map(header => header ? header.toString().trim() : '');

    // Handle duplicate column names
    const headerCount = {};
    const uniqueHeaders = headers.map((header) => {
        if (headerCount[header]) {
            headerCount[header] += 1;
            return `${header} (${headerCount[header]})`;
        } else {
            headerCount[header] = 1;
            return header;
        }
    });

    // Convert the sheet to JSON format, preserving empty cells and using unique headers
    const jsonData = xlsx.utils.sheet_to_json(sheet, { header: uniqueHeaders, defval: null, range: 1 });

    // Log the unique headers and sample data rows
    console.log('Unique Excel Headers:', uniqueHeaders);
    console.log('Sample data rows:', jsonData.slice(0, 3));

    return { workbook, sheetName, sheet, headers: uniqueHeaders, data: jsonData };
}




function mapMondayDataToExcel(headers, mondayItems) {
    const dataRows = [];

    mondayItems.forEach((item) => {
        const rowData = {};
        rowData['Year'] = item.name.trim(); // Assuming 'Year' is the item name

        item.column_values.forEach((colVal) => {
            const columnTitle = colVal.column ? colVal.column.title.trim() : '';
            const mondayValue = colVal.text ? colVal.text.trim() : '';

            // Exclude "Comment" column and ensure "Protech Generic System Name" data is captured
            if (columnTitle && columnTitle.toLowerCase() !== 'comment') {
                if (headers.includes(columnTitle)) {
                    rowData[columnTitle] = mondayValue !== 'No Text' ? mondayValue : '';
                } else {
                    console.warn(`Column "${columnTitle}" not found in Excel headers.`);
                }
            }
        });

        dataRows.push(rowData);
    });

    return dataRows;
}



function updateExcelSheetWithData(
    workbook,
    sheetName,
    sheet,
    headers,
    dataRows,
    filePath
) {
    // Get the range of the sheet
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const startRow = 1; // Assuming headers are in the first row (index 0)

    // Clear existing data rows starting from row 2
    for (let R = startRow + 1; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
            const cell = sheet[cellAddress];
            if (cell) {
                // Clear cell value but keep formatting
                cell.v = '';
            }
        }
    }

    // Write new data rows
    let currentRow = startRow + 1; // Start from the row after headers

    dataRows.forEach((rowData) => {
        headers.forEach((header, colIndex) => {
            const cellAddress = xlsx.utils.encode_cell({ r: currentRow, c: colIndex });
            let cell = sheet[cellAddress];

            if (!cell) {
                // If the cell doesn't exist, create it
                cell = { t: 's', v: '', s: {} };
            }

            // Update cell value
            const cellValue = rowData[header];
            cell.v = cellValue;

            // Save cell back to sheet
            sheet[cellAddress] = cell;
        });

        currentRow += 1;
    });

    // Update the sheet's used range
    sheet['!ref'] = xlsx.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: currentRow - 1, c: headers.length - 1 },
    });

    // Write the workbook back to file
    xlsx.writeFile(workbook, filePath);
}


// Implement the syncMondayToExcel function
async function syncMondayToExcel(boardId, apiKey, filePath) {
    const groupName = extractManufacturer(filePath); // Extract the group name from the Excel file name
    console.log(`Group Name extracted from file: ${groupName}`);

    // Read the existing Excel file with formatting
    console.log('Reading existing Excel file...');
    const { workbook, sheetName, sheet, headers, data } = readExcelFileWithFormatting(filePath);

    if (!Array.isArray(data) || data.length === 0) {
        console.warn("Excel file is empty or not properly formatted. Proceeding with synchronization.");
    }

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

    // Map Monday.com data to Excel columns
    console.log('Mapping Monday.com data to Excel format...');
    const mappedMondayData = mapMondayDataToExcel(headers, mondayItems);

    // Compare and update Excel data
    console.log('Comparing data and preparing updates...');
    const updatedData = compareAndUpdateExcelData(headers, data, mappedMondayData);

    // Write updated data back to the Excel file
    console.log('Writing updated data back to Excel file...');
    writeDataToExcelFile(workbook, sheetName, sheet, headers, updatedData, filePath);

    console.log(`Excel file updated successfully at ${filePath}`);
}



//########################################################################################################################


async function updateOrCreateBoard(filePath, apiKey, boardId) {
    // Read the existing Excel file with formatting
    const { workbook, sheetName, sheet, headers, data } = readExcelFileWithFormatting(filePath);
    const groupName = extractManufacturer(filePath);

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

    // Use headers directly
    const excelColumns = headers.map((col) => col.trim());

    console.log('Existing columns:', existingColumnTitles);

    // No need to create the "Year" column as it's the default item name column
    const columnIdMap = {};
    console.log('Skipping "Year" column creation as it is the item name column.');

    const startIndex = excelColumns[0].toLowerCase() === 'year' ? 1 : 0;

    // Create or map the remaining columns from the Excel sheet
    for (let i = startIndex; i < excelColumns.length; i++) {
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

    // Delete columns that exist on the board but are not in the Excel sheet, excluding 'Comment' column
    for (let existingColumn of existingColumns) {
        if (
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
    for (const row of data) {
        const year = String(row.Year).trim(); // Ensure 'Year' is used as the item name (default column)
        const columnValues = {};

        // Map Excel data to Monday.com columns using the correct column IDs, skipping the 'Year' column
        for (let i = startIndex; i < excelColumns.length; i++) {
            const column = excelColumns[i];
            const cellValue = row[column];

            // Check if the cell is not blank
            if (cellValue !== undefined && cellValue !== null && String(cellValue).trim() !== '') {
                columnValues[columnIdMap[column]] = String(cellValue).trim();
            } else {
                // Skip blanks
                console.log(`Skipping blank cell for column "${column}" in year "${year}"`);
            }
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
        if (Object.keys(columnValues).length > 0) {
            console.log(`Updating item ID ${newItemId} with column values.`);
            await updateItem(boardId, newItemId, columnValues, apiKey);
            console.log(`Updated item ID ${newItemId} with column values.`);
        } else {
            console.log(`No column values to update for item ID ${newItemId}. Skipping update.`);
        }
    }
}