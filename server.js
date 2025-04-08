const express = require('express');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs');
const path = require('path');
const os = require('os');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(express.static('public'));
app.use(express.json());
app.use(fileUpload());

// Store last data in memory
let lastData = [];

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/step4', (req, res) => {
    res.sendFile(path.join(__dirname, 'step4', 'index.html'));
});

app.get('/step5', (req, res) => {
    res.sendFile(path.join(__dirname, 'step5', 'index.html'));
});

app.use('/resources', express.static(path.join(__dirname, 'resources')));

app.post('/upload', async (req, res) => {
    try {
        if (!req.files || !req.files.file) {
            return res.status(400).send('No file uploaded');
        }

        const file = req.files.file;
        const tempPath = path.join(os.tmpdir(), file.name);
        
        // Save file temporarily
        await file.mv(tempPath);

        // Read Excel file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(tempPath);
        const worksheet = workbook.worksheets[0];

        // Process data (starting from row 2 for headers)
        const data = [];
        let headers = {};

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 2) { // Get headers from row 2
                row.eachCell((cell, colNumber) => {
                    headers[colNumber] = cell.text || cell.value || '';
                });
            } else if (rowNumber > 2) { // Process data rows
                const rowData = {};
                row.eachCell((cell, colNumber) => {
                    rowData[headers[colNumber]] = cell.text || cell.value || '';
                });

                // Only add rows that have at least some data and valid Market value
                if (Object.values(rowData).some(value => value) && 
                    ['DACH', 'France', 'Spain', 'Italy'].includes(rowData.Market)) {
                    data.push(rowData);
                }
            }
        });

        // Get last 25 records
        lastData = data.slice(-25);

        // Clean up temp file
        fs.unlinkSync(tempPath);

        res.json(lastData);
    } catch (error) {
        console.error(error);
        res.status(500).send('Error processing file');
    }
});

app.post('/process', async (req, res) => {
    try {
        const selectedIndices = req.body.rows.map(Number);
        const step = req.body.step || 'step4';

        // Validate indices
        if (selectedIndices.some(idx => idx < 0 || idx >= lastData.length)) {
            return res.status(400).send('Invalid selection');
        }

        // Create new workbook and copy template to temp location
        const workbook = new ExcelJS.Workbook();
        const templateName = step === 'step5' ? 'PlantillaSTEP5.xlsx' : 'PlantillaSTEP4.xlsx';
        const templatePath = path.join(os.tmpdir(), `template_${Date.now()}_${templateName}`);
        
        // Copy template file to temp location
        const templateSourcePath = path.join(__dirname, 'templates', templateName);
        if (!fs.existsSync(templateSourcePath)) {
            return res.status(500).send(`Template file ${templateName} not found`);
        }
        fs.copyFileSync(templateSourcePath, templatePath);
        
        try {
            await workbook.xlsx.readFile(templatePath);
            const worksheet = workbook.worksheets[0];
            
            let destinationRow = step === 'step5' ? 2 : 7;

            for (const idx of selectedIndices) {
            const record = lastData[idx];
            const { Name, Surname, 'E-mail': email, Market, 'Va a ser PCC?': pccStatus, 'B2E User Name': b2eUsername } = record;

            if (step === 'step5') {
                // Step 5 logic - fill the first sheet with queue information
                const windowsUserName = record['Windows User Name'] || ''; // Get Windows User Name from column N
                worksheet.getCell(`A${destinationRow}`).value = windowsUserName ? `${windowsUserName}.pcc@costa.it` : '';
                // Set Queue Name as uppercase combination of Name and Surname
                let marketQueue = '';
                switch (Market) {
                    case 'DACH':
                        marketQueue = 'Costa Kreuzfahrten';
                        break;
                    case 'France':
                        marketQueue = 'Costa Croisieres';
                        break;
                    case 'Italy':
                        marketQueue = 'Costa Crociere';
                        break;
                    case 'Spain':
                        marketQueue = 'Costa Cruceros';
                        break;
                }
                const queueName = `${Name} ${Surname} - ${marketQueue}`.toUpperCase();
                worksheet.getCell(`B${destinationRow}`).value = queueName;
                // Fill Queue Owner (Column C)
                worksheet.getCell(`C${destinationRow}`).value = `${Name} ${Surname}`;
                // Fill SLA (Column D)
                worksheet.getCell(`D${destinationRow}`).value = '2 days';
                // Fill User to be added (Column E)
                worksheet.getCell(`E${destinationRow}`).value = `${Name} ${Surname}`;
                // Fill Team Leader (Column F)
                let teamLeader = '';
                if (pccStatus === 'DS') {
                    teamLeader = 'Valeria Amoroso';
                } else if (Market === 'DACH' || (Market === 'France' && pccStatus !== 'N')) {
                    teamLeader = 'Adrian Arias';
                } else {
                    teamLeader = 'Valeria Amoroso';
                }
                worksheet.getCell(`F${destinationRow}`).value = teamLeader;
                destinationRow++;

                // Process Foglio1 sheet
                const foglio1 = workbook.getWorksheet('Foglio1');
                if (foglio1) {
                    // Add data to Foglio1 sheet starting from row 2
                    const foglio1Row = destinationRow - 1;
                    foglio1.getCell(`A${foglio1Row}`).value = `${Name} ${Surname}`; // PCC column
                    foglio1.getCell(`B${foglio1Row}`).value = email; // Outlook column
                    foglio1.getCell(`C${foglio1Row}`).value = windowsUserName ? `${windowsUserName}.pcc@costa.it` : ''; // D365 column
                }
                continue;
            }

            // Fill columns C, D, E
            worksheet.getCell(`C${destinationRow}`).value = Name;
            worksheet.getCell(`D${destinationRow}`).value = Surname;
            worksheet.getCell(`E${destinationRow}`).value = email;

            // Fill column F (Primary phone)
            if (pccStatus === 'Y') {
                if (Market === 'DACH') {
                    worksheet.getCell(`F${destinationRow}`).value = '+4940210918145';
                } else if (Market === 'France') {
                    worksheet.getCell(`F${destinationRow}`).value = '+33180037979';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`F${destinationRow}`).value = '+34932952130';
                } else if (Market === 'Italy') {
                    worksheet.getCell(`F${destinationRow}`).value = '+390109997099';
                }
            } else {
                worksheet.getCell(`F${destinationRow}`).value = '';
            }

            // Fill column G (Workgroup)
            if (pccStatus === 'Y' || pccStatus === 'TL') {
                if (Market === 'DACH') {
                    worksheet.getCell(`G${destinationRow}`).value = 'D_PCC';
                } else if (Market === 'France') {
                    worksheet.getCell(`G${destinationRow}`).value = 'F_PCC';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`G${destinationRow}`).value = 'E_PCC';
                } else if (Market === 'Italy') {
                    worksheet.getCell(`G${destinationRow}`).value = 'I_PCC';
                }
            } else if (pccStatus === 'N' || pccStatus === 'DS') {
                if (Market === 'DACH') {
                    worksheet.getCell(`G${destinationRow}`).value = 'D_Outbound';
                } else if (Market === 'France') {
                    worksheet.getCell(`G${destinationRow}`).value = 'F_Outbound';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`G${destinationRow}`).value = 'E_Outbound';
                }
            }

            // Fill column H (Workgroup WA)
            if (pccStatus === 'Y' || pccStatus === 'DS') {
                if (Market === 'DACH') {
                    worksheet.getCell(`H${destinationRow}`).value = 'D_WAPCC';
                } else if (Market === 'France') {
                    worksheet.getCell(`H${destinationRow}`).value = 'F_WAPCC';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`H${destinationRow}`).value = 'E_WAPCC';
                } else if (Market === 'Italy') {
                    worksheet.getCell(`H${destinationRow}`).value = 'I_WAPCC';
                }
            }

            // Fill column I (Team)
            if (pccStatus === 'Y' || pccStatus === 'TL') {
                if (Market === 'DACH') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_D_CCH_PCC_1';
                } else if (Market === 'France') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_F_CCH_PCC_1';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_E_CCH_PCC_1';
                } else if (Market === 'Italy') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_I_CCH_PCC_1';
                }
            } else if (pccStatus === 'N' || pccStatus === 'DS') {
                if (Market === 'DACH') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_D_CCH_B2C_1';
                } else if (Market === 'France') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_F_CCH_B2C_1';
                } else if (Market === 'Spain') {
                    worksheet.getCell(`I${destinationRow}`).value = 'Team_E_CCH_B2C_1';
                }
            }

            // Fill column M (Is PCC)
            worksheet.getCell(`M${destinationRow}`).value = pccStatus === 'Y' ? 'Y' : 'N';

            // Fill columns Q, R, S (CCRM, CTI User y TTG UserID 1)
            worksheet.getCell(`Q${destinationRow}`).value = email;
            worksheet.getCell(`R${destinationRow}`).value = email;
            worksheet.getCell(`S${destinationRow}`).value = b2eUsername;

            // Fill column W (Campaign Level)
            if (pccStatus === 'TL') {
                worksheet.getCell(`W${destinationRow}`).value = 'Team Leader';
            } else {
                worksheet.getCell(`W${destinationRow}`).value = 'Agent';
            }

            destinationRow++;
        }

        // Save file with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '');
        const outputPath = path.join(os.tmpdir(), `${step === 'step5' ? 'D365_STEP5_CCH' : 'D365_STEP4_CCH'}_${timestamp}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);

        res.download(outputPath, `${step === 'step5' ? 'D365_STEP5_CCH' : 'D365_STEP4_CCH'}_${timestamp}.xlsx`, (err) => {
            if (err) console.error(err);
            // Clean up temp file after download
            fs.unlinkSync(outputPath);
            });
        } catch (error) {
            console.error(error);
            res.status(500).send('Error processing request');
        }
    } catch (error) {
        console.error(error);
        res.status(500).send('Error processing request');
    }
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});