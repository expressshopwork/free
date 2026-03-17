// Import packages
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');

// Initialize app
const app = express();
app.use(cors());
app.use(bodyParser.json());

// GS_URL changed here
const GS_URL = 'https://script.google.com/macros/s/AKfycbwv8kdASEnPxjQJ8-wvnIZ5_Ib7_suDDcYU46VPNeVE0WKJboAYQlq3TqrA1QsXu0Y/exec';

// Define routes
app.get('/', (req, res) => {
    res.send('Hello World!');
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});