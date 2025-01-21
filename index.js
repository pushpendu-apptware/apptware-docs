// setup a express server
const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const config = require('./config.json');
const dotenv = require('dotenv');

// Add these for better error handling
const InspectModule = require("docxtemplater/js/inspect-module");
const expressions = require("docxtemplater/js/expressions.js");

dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const host = process.env.HOST || 'localhost';

app.use(express.json());

// Add static file serving for downloads
app.use('/downloads', express.static(path.join(__dirname, 'output')));

app.get('/', (req, res) => {
    res.send('Hello World!');
});

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
        const downloadUrl = `http://${host}:${port}/downloads/${fileName}`;

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

// Add download endpoint
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'output', req.params.filename);
    res.download(filePath);
});

app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`);
});