const express = require('express');
const path = require('path');
const app = express();
const port = 3000;

// Serve static files from the 'public' directory
app.use(express.static('public'));

// You can also serve from multiple directories
app.use(express.static('images'));
app.use(express.static('css'));

// Or with a virtual path prefix
app.use('/static', express.static('public'));

// Route for the homepage
app.get('/', (req, res) => {
    res.send('Hello World!');
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});