import express from 'express';
import bodyParser from 'body-parser';
import dataJson from './response.json';

const app = express();
const Excel = require('exceljs');

app.use(bodyParser.urlencoded({ extended: true }));

app.get('/', (req, res, next) => {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('ExampleSheet');

  dataJson.map((data, key) => {
    let row = worksheet.getRow(key + 2);
    let header = worksheet.getRow(1);
    Object.keys(data).map((valueObj, i) => {
      if (valueObj !== 'items') {
        // console.log(Object.keys(data).length);
        // row.getCell(i + 1).value = data[valueObj];
        header.getCell(String.fromCharCode(65 + i)).value = valueObj;
        row.getCell(String.fromCharCode(65 + i)).value = data[valueObj];
      } else {
        data.items.map((itemValue, key) => {
          Object.keys(itemValue).map((itemObj, index) => {
            //menentukan nama header column

            if (itemObj === 'billTypeID') {
              header.getCell(
                Object.keys(itemValue).length * key + Object.keys(data).length + index,
              ).value = `item-${key}-name`;
            } else {
              header.getCell(
                Object.keys(itemValue).length * key + Object.keys(data).length + index,
              ).value = `item-${key}-${itemObj}`;
            }

            if (itemObj !== 'subBillType') {
              if (itemObj === 'billTypeID') {
                row.getCell(
                  Object.keys(itemValue).length * key + Object.keys(data).length + index,
                ).value = itemValue[itemObj]['name'];
              } else {
                row.getCell(
                  Object.keys(itemValue).length * key + Object.keys(data).length + index,
                ).value = itemValue[itemObj];
              }
            } else {
              row.getCell(
                Object.keys(itemValue).length * key + Object.keys(data).length + index,
              ).value = itemValue[itemObj];

              // itemValue.subBillType.map(subItemValue => {
              //   Object.keys(subItemValue).map((itemSubObj, indexSub) => {
              //     row.getCell(
              //       Object.keys(data).length + Object.keys(itemValue).length + indexSub - 1,
              //     ).value = subItemValue[itemSubObj];
              //   });
              // });
            }
          });
        });
      }
    });
  });

  // worksheet.addRow(rowValues);
  // worksheet.addRow(rowValues);

  //menambahkan row
  // worksheet.addRows(rows);

  // save workbook to disk
  workbook.xlsx
    .writeFile('sample.xlsx')
    .then(() => {
      console.log('saved');
    })
    .catch(err => {
      console.log('err', err);
    });

  res.jsonp('index');
});

app.post('/qrcode', async (req, res, next) => {
  let data = { name: 'hari irawan' };
  res.jsonp(data);
});

app.listen(1994, () => console.log('Server started at port 1994'));
