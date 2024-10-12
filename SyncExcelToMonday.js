const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');

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
    const response = await axios.post('https://api.monday.com/v2', { query, variables }, {
      headers: { Authorization: apiKey }
    });

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
    groupTitle
  };

  try {
    const response = await axios.post('https://api.monday.com/v2', { query: mutation, variables }, {
      headers: { Authorization: apiKey }
    });

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
    const response = await axios.post('https://api.monday.com/v2', { query, variables }, {
      headers: { Authorization: apiKey }
    });

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
    columnType: "text"
  };

  try {
    const response = await axios.post('https://api.monday.com/v2', { query: mutation, variables }, {
      headers: { Authorization: apiKey }
    });

    if (response.data.errors) {
      console.error('GraphQL errors:', response.data.errors);
      return null;
    }

    return response.data.data.create_column.id;
  } catch (error) {
    console.error('Error creating column:', error);
  }
}

// Function to create a new item in a specific group on the board (with year as item name)
async function createItemInGroup(boardId, groupId, itemName, columnValues, apiKey) {
  const mutation = `
    mutation($boardId: ID!, $groupId: String!, $itemName: String!) {
      create_item(board_id: $boardId, group_id: $groupId, item_name: $itemName) {
        id
      }
    }
  `;

  const variables = {
    boardId: boardId,  // Ensure this is passed as an integer
    groupId: groupId,  // Ensure this is passed as a string
    itemName: itemName  // The Year from Excel, this is the "item name" for the first slot
  };

  try {
    const response = await axios.post('https://api.monday.com/v2', { query: mutation, variables }, {
      headers: { Authorization: apiKey }
    });

    if (response.data.errors) {
      console.error('GraphQL errors:', response.data.errors);
      return null;
    }

    if (response.data && response.data.data && response.data.data.create_item) {
      return response.data.data.create_item.id;  // Return the new item's ID
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
    boardId: boardId,  // Ensure this is passed as an integer
    itemId: itemId,    // The newly created item's ID
    columnValues: JSON.stringify(columnValues)  // Convert the column values to JSON
  };

  try {
    const response = await axios.post('https://api.monday.com/v2', { query: mutation, variables }, {
      headers: { Authorization: apiKey }
    });

    if (response.data.errors) {
      console.error('GraphQL errors:', response.data.errors);
      return null;
    }

    if (response.data && response.data.data && response.data.data.change_multiple_column_values) {
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
    columnId: columnId
  };

  try {
    const response = await axios.post('https://api.monday.com/v2', { query: mutation, variables }, {
      headers: { Authorization: apiKey }
    });

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

// Function to fetch existing items in a specific group on the board
async function fetchItemsInGroup(boardId, groupId, apiKey) {
    const query = `
    query ($boardId: [Int], $groupId: String!) {
      boards (ids: $boardId) {
        groups(ids: [$groupId]) {
          id
          items {
            id
            name
            column_values {
              id
              title
              text
            }
          }
        }
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
            { query, variables },
            {
                headers: { Authorization: apiKey },
            }
        );

        if (response.data.errors) {
            console.error('GraphQL errors:', response.data.errors);
            return [];
        }

        return response.data.data.boards[0].groups[0].items;
    } catch (error) {
        console.error('Error fetching items in group:', error);
        return [];
    }
}


// Main function to update or create board based on Excel file and ensure columns match
async function updateOrCreateBoard(filePath, apiKey, boardId) {
    const excelData = readExcelFile(filePath);
    const groupName = extractManufacturer(filePath); // Use the first word of the file name as the group title

    // Fetch existing groups on the board
    const groups = await fetchGroups(boardId, apiKey);

    // Check if the group already exists or create a new group
    let groupId;
    const existingGroup = groups.find(
        (group) =>
            group.title.trim().toLowerCase() === groupName.trim().toLowerCase()
    );
    if (existingGroup) {
        groupId = existingGroup.id;
        console.log(`Group "${groupName}" already exists with ID: ${groupId}`);
    } else {
        groupId = await createGroup(boardId, groupName, apiKey);
        if (!groupId) {
            console.error("Failed to create group.");
            return;
        }
        console.log(`Created new group: ${groupName} with ID: ${groupId}`);
    }

    // Fetch existing columns on the board
    const existingColumns = await fetchColumns(boardId, apiKey);
    const existingColumnTitles = existingColumns.map((col) =>
        col.title.trim().toLowerCase()
    );

    console.log("Existing columns:", existingColumnTitles);

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

    // Delete columns that exist on the board but are not in the Excel sheet
    for (let existingColumn of existingColumns) {
        if (
            existingColumn.title.trim().toLowerCase() !== 'name' && // Do not delete the 'Name' column
            !excelColumns.some(
                (excelCol) =>
                    excelCol.toLowerCase() === existingColumn.title.trim().toLowerCase()
            )
        ) {
            console.log(`Deleting extra column: ${existingColumn.title}`);
            await deleteColumn(boardId, existingColumn.id, apiKey);
        }
    }

    // Fetch existing items in the group
    const existingItems = await fetchItemsInGroup(boardId, groupId, apiKey);
    const existingItemsMap = {}; // Map of item name (year) to item data

    existingItems.forEach((item) => {
        const itemName = item.name.trim();
        const columnValues = {};
        item.column_values.forEach((colVal) => {
            columnValues[colVal.title.trim()] = colVal.text.trim();
        });
        existingItemsMap[itemName] = {
            id: item.id,
            column_values: columnValues,
        };
    });

    // Iterate over Excel data and create or update items
    for (const row of excelData) {
        const year = String(row.Year).trim(); // Ensure 'Year' is used as the item name (default column)
        const columnValues = {};

        // Map Excel data to Monday.com columns using the correct column IDs, skipping the 'Year' column
        for (let i = 1; i < excelColumns.length; i++) {
            const column = excelColumns[i].trim();
            columnValues[columnIdMap[column]] = String(row[column]).trim();
        }

        if (existingItemsMap[year]) {
            // Item exists, compare data
            const existingItem = existingItemsMap[year];
            const existingColumnValues = existingItem.column_values;

            let isDifferent = false;
            for (let i = 1; i < excelColumns.length; i++) {
                const column = excelColumns[i].trim();
                const excelValue = String(row[column]).trim();
                const existingValue = existingColumnValues[column] || '';

                if (excelValue !== existingValue) {
                    isDifferent = true;
                    break;
                }
            }

            if (isDifferent) {
                console.log(`Updating existing item for year ${year}`);
                await updateItem(boardId, existingItem.id, columnValues, apiKey);
                console.log(`Updated item ID ${existingItem.id} with new values.`);
            } else {
                console.log(`No changes for item year ${year}, skipping update.`);
            }
        } else {
            // Item does not exist, create it
            console.log(`Creating new item for year ${year}`);
            const newItemId = await createItemInGroup(
                boardId,
                groupId,
                year,
                columnValues,
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
}

// Call the function with your Excel file path, API key, and board ID
updateOrCreateBoard(
  'Acura Pre-Qual Long Sheet v6.3.xlsx',
  'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQwNzM1MzIxNywiYWFpIjoxMSwidWlkIjo0MTI5ODM0MCwiaWFkIjoiMjAyNC0wOS0wNlQxNjo0MjozMi4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTIzOTkzMjYsInJnbiI6InVzZTEifQ._QYJKxEcmmUB6-en7MKIPHXw3s-7_lNGDVFBLjNjK18', // Replace with your actual API key
  '7620836068' // Replace with your actual board ID
);