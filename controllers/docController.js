const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const InspectModule = require("docxtemplater/js/inspect-module");

const config = require('../config.json');

exports.generateDoc = (req, res) => {
    try {
        const selectedDocs = [];
        console.log('Request body:', req.body);
        if (req.body.nda) selectedDocs.push("nda");
        if (req.body.msa) selectedDocs.push("msa");

        if (selectedDocs.length === 0) {
            throw new Error("No document selected");
        }

        const downloadUrls = [];

        selectedDocs.forEach((docKey) => {
            const docConfig = config[docKey];
            // Process required input variables from config, skip auto-generated companyTitle
            const templateVars = docConfig.inputs.reduce((acc, input) => {
                const varName = input.alias.replace('$', '');
                if(varName === 'companyTitle'){
                    return acc;
                }
                if (!req.body[varName]) {
                    throw new Error(`Missing required variable: ${input.name}`);
                }
                acc[varName] = req.body[varName];
                return acc;
            }, {});

            // Automatically generate company title from company address.
            if (!req.body.companyAddress) {
                throw new Error("Missing required variable: Company Address");
            }
            // Generate companyTitle from companyName (using the first word)
            const computedTitle = req.body.companyName.split(' ')[0];
            templateVars.companyTitle = computedTitle;

            // Read the template and generate document using InspectModule for debugging
            const template = fs.readFileSync(docConfig.path, 'binary');
            const zip = new PizZip(template);
            const iModule = InspectModule();
            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                modules: [iModule]
            });
            doc.render(templateVars);

            console.log('Tags found in template:', iModule.getAllTags());

            // Determine file name
            let fileName;
            if (req.body.companyName) {
                const safeCompanyName = req.body.companyName
                    .replace(/[^a-zA-Z0-9]/g, '-')
                    .replace(/-+/g, '-')
                    .replace(/(^-|-$)/g, '');
                fileName = `${docKey}-${safeCompanyName}-${Date.now()}.docx`;
            } else {
                fileName = `${docKey}-${Date.now()}.docx`;
            }

            const outputDir = path.join(__dirname, '../output');
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir);
            }
            const outputPath = path.join(outputDir, fileName);
            fs.writeFileSync(outputPath, doc.getZip().generate({ type: 'nodebuffer' }));

            const host = process.env.HOST || 'http://localhost:3000';
            downloadUrls.push(`${host}/downloads/${fileName}`);

            // once the download urls got generated automatically giving get requests to that urls
            // to download the files


        });

        res.json({ success: true, downloadUrls });
    } catch (error) {
        console.error('Template error:', error);
        res.status(500).json({
            error: error.message,
            requiredVariables: config.nda.inputs
                .filter(i => i.alias.replace('$', '') !== 'companyTitle')
                .map(i => ({
                    name: i.name,
                    variable: i.alias.replace('$', '')
                }))
                .concat({
                    name: 'Company Address',
                    variable: 'companyAddress'
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
