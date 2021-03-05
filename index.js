const express = require('express');
const cors = require('cors')
const app = express();
const port = 3000;

const ExcelJS = require('exceljs');

app.use(cors());

app.use(express.json({limit: '200mb'}));
app.use(express.urlencoded({limit: '200mb', extended: true}));

async function createWorkbook(json) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Plan1');
    
    Object.keys(json)
        .filter(key => key != 'rows')
        .forEach(key => sheet[key] = json[key]);

    json.rows.forEach((rowOptions, index) => {
        const row = sheet.getRow(parseInt(index) + 1);

        Object.keys(rowOptions)
            .filter(key => key != 'cells')
            .forEach(key => row[key] = rowOptions[key]);

        rowOptions.cells.forEach(cellOptions => {
            const refs = cellOptions.ref.split(':');

            if(refs.length > 0) {
                sheet.mergeCells(cellOptions.ref);
            }

            const cell = sheet.getCell(refs[0]);

            //borda padrao
            cell.border = {
                top:    {style:'thin'},
                left:   {style:'thin'},
                bottom: {style:'thin'},
                right:  {style:'thin'}
            };

            Object.keys(cellOptions)
                .filter(key => key != 'ref')
                .forEach(key => cell[key] = cellOptions[key]);
            
        });
    });

    return workbook;
};

app.post('/api/plan', async (req, res) => {
    try {
        const workbook   = await createWorkbook(req.body);
        const fileBuffer = await workbook.xlsx.writeBuffer();

        res.setHeader('Content-Type' ,'text/plain');
        res.status(200).send(fileBuffer.toString('base64'));
    } catch(error) {
        res.status(400).json({ error: error.toString() });
    }
})

app.listen(port, () => console.log(`Listening at port: ${port}`));