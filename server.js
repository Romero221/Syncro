// JavaScript source code

const express = require('express'); // Import the Express library
const app = express(); // Create an instance of an Express application
const PORT = 8080; // Define the port the server will listen on

// Serve static files (like HTML, CSS, JS) from the 'public' directory
app.use(express.static('public'));

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
