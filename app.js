const express = require('express');
const app = express();

// SmartNas item definition
const items = [
    { id: 'i1', name: 'Item 1', shortcut: 'I1' },
    { id: 'i2', name: 'Item 2', shortcut: 'I2' },
    { id: 'i3', name: 'Item 3', shortcut: 'I3' },
    { id: 'i4', name: 'SmartNas', shortcut: 'SNA' },  // Changed from 'SN' to 'SNA'
    // other items...
];

app.get('/items', (req, res) => {
    res.json(items);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
