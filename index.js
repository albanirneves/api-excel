const express = require('express');
const cors = require('cors')
const app = express();
const port = 3000;

const ExcelJS  = require('exceljs');
const colCache = require('exceljs/lib/utils/col-cache');
const axios    = require('axios');

app.use(cors());

app.use(express.json({limit: '200mb'}));
app.use(express.urlencoded({limit: '200mb', extended: true}));

async function createWorkbook(json) {
    const workbook = new ExcelJS.Workbook();

    //seta as propriedades do workboot
    workbook.calcProperties.fullCalcOnLoad = true;

    //seta as propriedades da planilha (tab)
    const worksheet = workbook.addWorksheet('Plan1');
    
    Object.keys(json)
        .filter(key => key != 'rows' && key != 'workbook')
        .forEach(key => worksheet[key] = json[key]);

    //seta as propriedades de cada linha
    await Promise.all(json.rows.map(async (rowOptions, index) => {
        const row = worksheet.getRow(parseInt(index) + 1);

        Object.keys(rowOptions)
            .filter(key => key != 'cells')
            .forEach(key => row[key] = rowOptions[key]);

        //seta as propriedades de cada célula
        await Promise.all(rowOptions.cells.map(async cellOptions => {
            const refs = cellOptions.ref.split(':');
            const cell = worksheet.getCell(refs[0]);
            
            if(cellOptions.image) {
                let imgId;

                if(cellOptions.image.link.slice(0, 4) == 'http') {
                    const response = await axios.get(cellOptions.image.link,  { responseType: 'arraybuffer' })
                    const buffer = Buffer.from(response.data, 'utf-8');
    
                    imgId = workbook.addImage({
                        buffer: buffer,
                        extension: cellOptions.image.link.split('.').pop(),
                    });
    
                } else if(cellOptions.image.link.slice(0, 4) == 'data') {
                    imgId = workbook.addImage({
                        base64: cellOptions.image.link,
                        extension: cellOptions.image.link.split(';')[0].split('/')[1],
                    });
                }

                if(cellOptions.image.width && cellOptions.image.height) {
                    const { col,  row }    = colCache.decode(refs[0]);
                    const { width, height } = cellOptions.image

                    worksheet.addImage(imgId, {
                        tl:  { col: col - 1, row: row - 1 },
                        ext: { width, height }
                    });
                } else {
                    worksheet.addImage(imgId, cellOptions.ref);
                }

            } else {
                if(refs.length > 0) {
                    worksheet.mergeCells(cellOptions.ref);
                }

                Object.keys(cellOptions)
                    .filter(key => key != 'ref' && key != 'image')
                    .forEach(key => cell[key] = cellOptions[key]);
            }
        }));
    }));

    return workbook;
};

app.post('/api/plan', async (req, res) => {
    try {
        const workbook   = await createWorkbook(req.body);
        const fileBuffer = await workbook.xlsx.writeBuffer();

        res.setHeader('Content-Type' ,'text/plain');
        res.status(200).send(fileBuffer.toString('base64'));
    } catch(error) {
        res.status(400).json({ error: error.stack });
    }
})

app.listen(port, () => console.log(`Listening at port: ${port}`));