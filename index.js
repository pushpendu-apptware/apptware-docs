const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const path = require('path');

dotenv.config();
const app = express();
const port = process.env.PORT || 3000;

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'DELETE', 'UPDATE', 'PUT', 'PATCH', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin'],
    credentials: true
}));

app.use(express.json());

// Serve static files for downloads
app.use('/downloads', express.static(path.join(__dirname, 'output')));

// Use routes
const docRoutes = require('./routes/docRoutes');
app.use('/', docRoutes);

app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`);
});