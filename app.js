const express = require('express');
const multer = require('multer');
const exceljs = require('exceljs');
const mongodb = require('mongodb');

const app = express();
const port = 3000;

// Multer configuration
const upload = multer({ dest: 'uploads/' });

// MongoDB connection
const MongoClient = mongodb.MongoClient;
const mongoUrl = 'mongodb://localhost:27017';
const dbName = 'gpfflab_4';

// Route to handle file upload
app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        // Connect to MongoDB
        const client = await MongoClient.connect(mongoUrl);
        const db = client.db(dbName);

        // Read Excel file
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile(req.file.path);

        // Get the first worksheet
        const worksheet = workbook.getWorksheet(1);

        // Iterate over each row and insert into MongoDB
        // const collection = db.collection('volunteers_new');
        // worksheet.eachRow((row, rowNumber) => {
        //     if (rowNumber !== 1) { // Skip header row
        //         const rowData = row.values.slice(1); // Skip first cell
        //         collection.insertOne({ data: rowData });
        //     }
        // });

        //new
        const collection = db.collection('volunteers_new');
        worksheet.eachRow(async (row, rowNumber) => {
            if (rowNumber !== 1) { // Skip header row
                try {
                    // Convert row values to an object
                    const rowData = {};
                    row.eachCell((cell, colNumber) => {
                        // Assuming column headers are in the first row
                        const header = worksheet.getRow(1).getCell(colNumber).value;
                        rowData[header] = cell.value;
                    });

                    // Insert row data into MongoDB
                    await collection.insertOne(rowData);
                } catch (err) {
                    console.error(`Error inserting row ${rowNumber}: ${err}`);
                }
            }
        });

        res.send('Data imported successfully');
    } catch (err) {
        console.error(err);
        res.status(500).send('Error importing data');
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server is listening at http://localhost:${port}`);
});
