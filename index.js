// setup a express server
const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const config = require('./config.json');
const dotenv = require('dotenv');
const cors = require('cors');

// Add these for better error handling
const InspectModule = require("docxtemplater/js/inspect-module");
const expressions = require("docxtemplater/js/expressions.js");

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const host = process.env.HOST || 'localhost';

// Configure CORS with all permissions
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'DELETE', 'UPDATE', 'PUT', 'PATCH', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin'],
    credentials: true
}));

app.use(express.json());

// Add static file serving for downloads
app.use('/downloads', express.static(path.join(__dirname, 'output')));

app.post('/generate-doc', (req, res) => {
    try {
        // Transform input data to match template variables (without $ prefix)
        const templateVars = config.nda.inputs.reduce((acc, input) => {
            const varName = input.alias.replace('$', '');
            if (!req.body[varName]) {
                throw new Error(`Missing required variable: ${input.name}`);
            }
            // Don't include $ in the key
            acc[varName] = req.body[varName];
            return acc;
        }, {});

        console.log('Template variables:', templateVars);

        const template = fs.readFileSync(config.nda.path, 'binary');
        const zip = new PizZip(template);

        // Add inspect module for debugging
        const iModule = InspectModule();

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            modules: [iModule]
        });

        // Render with clean variable names
        doc.render(templateVars);

        // Log tags that were found in the template
        console.log('Tags found in template:', iModule.getAllTags());

        const buffer = doc.getZip().generate({ type: 'nodebuffer' });

        // Get original filename without extension
        const originalName = path.basename(config.nda.path, '.docx');

        // Use companyName from request body for filename
        const safeCompanyName = req.body.companyName
            .replace(/[^a-zA-Z0-9]/g, '-')
            .replace(/-+/g, '-')
            .replace(/(^-|-$)/g, '');

        const fileName = `${originalName}-${safeCompanyName}-${Date.now()}.docx`;
        const outputPath = path.join(__dirname, 'output', fileName);

        // Create output directory if it doesn't exist
        if (!fs.existsSync(path.join(__dirname, 'output'))) {
            fs.mkdirSync(path.join(__dirname, 'output'));
        }

        fs.writeFileSync(outputPath, buffer);

        // Generate download URL
        const downloadUrl = `https://${host}/downloads/${fileName}`;

        res.json({
            success: true,
            file: outputPath,
            downloadUrl: downloadUrl
        });
    } catch (error) {
        console.log('Template error:', error);
        if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors.map(e => e.properties.explanation).join("\n");
            console.log('Detailed error:', errorMessages);
        }
        res.status(500).json({
            error: error.message,
            requiredVariables: config.nda.inputs.map(i => ({
                name: i.name,
                variable: i.alias.replace('$', '')
            }))
        });
    }
});

// Update download endpoint with stream handling
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'output', req.params.filename);

    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File not found' });
    }

    // Set headers
    res.setHeader('Content-Disposition', `attachment; filename=${req.params.filename}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    // Create read stream
    const fileStream = fs.createReadStream(filePath);

    // Handle stream events
    fileStream.on('error', (error) => {
        console.error('Stream error:', error);
        res.status(500).end();
    });

    // Pipe the file to response
    fileStream.pipe(res);

    // When response is finished, delete the file
    res.on('finish', () => {
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error('Error deleting file:', err);
            } else {
                console.log('Successfully deleted file:', filePath);
            }
        });
    });
});

app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`);
});