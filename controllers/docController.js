const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const InspectModule = require("docxtemplater/js/inspect-module");

const config = require('../config.json');

exports.generateDoc = (req, res) => {
    try {
        // Transform input data to match template variables (without $ prefix)
        const templateVars = config.nda.inputs.reduce((acc, input) => {
            const varName = input.alias.replace('$', '');
            if (!req.body[varName]) {
                throw new Error(`Missing required variable: ${input.name}`);
            }
            acc[varName] = req.body[varName];
            return acc;
        }, {});

        // Ensure companyTitle is provided, then add it to templateVars
        if (!req.body.companyTitle) {
            throw new Error('Missing required variable: Company Title');
        }
        templateVars['companyTitle'] = req.body.companyTitle;

        // Extract first name from companyAddress and add it to templateVars
        const companyAddressFirstName = req.body.companyAddress.split(' ')[0];
        templateVars['companyAddressFirstName'] = companyAddressFirstName;

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
        const outputPath = path.join(__dirname, '../output', fileName);

        // Create output directory if it doesn't exist
        if (!fs.existsSync(path.join(__dirname, '../output'))) {
            fs.mkdirSync(path.join(__dirname, '../output'));
        }

        fs.writeFileSync(outputPath, buffer);

        const host = process.env.HOST || 'localhost:3000';
        const downloadUrl = `http://${host}/downloads/${fileName}`;

        res.json({
            success: true,
            file: outputPath,
            downloadUrl: downloadUrl
        });
    } catch (error) {
        console.error('Template error:', error);
        if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors.map(e => e.properties.explanation).join("\n");
            console.log('Detailed error:', errorMessages);
        }
        res.status(500).json({
            error: error.message,
            requiredVariables: config.nda.inputs.map(i => ({
                name: i.name,
                variable: i.alias.replace('$', '')
            })).concat({
                name: 'Company Title',
                variable: 'companyTitle'
            })
        });
    }
};

exports.downloadDoc = (req, res) => {
    const filePath = path.join(__dirname, '../output', req.params.filename);

    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File not found' });
    }

    res.setHeader('Content-Disposition', `attachment; filename=${req.params.filename}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    const fileStream = fs.createReadStream(filePath);

    fileStream.on('error', (error) => {
        console.error('Stream error:', error);
        res.status(500).end();
    });

    fileStream.pipe(res);

    // Delete the file after response is finished
    res.on('finish', () => {
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error('Error deleting file:', err);
            } else {
                console.log('Successfully deleted file:', filePath);
            }
        });
    });
};

exports.homePage = (req, res) => {
    res.send('Hello World!');
};