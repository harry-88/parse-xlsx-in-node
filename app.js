const http = require('http');

const hostname = '127.0.0.1';
const fs = require('fs');
const XLSX = require('xlsx')
const port = 3000;


const getRecord = () => {

    const file = XLSX.readFile('D://parse-xlsx-in-node/empty.xls');
    let data = []
    const sheets = file.SheetNames
    for (let i = 0; i < sheets.length; i++) {
        const temp = XLSX.utils.sheet_to_json(
            file.Sheets[file.SheetNames[i]])
        data = temp;
    }
    let record = [];
    let flag = 0;
    for (let i = 0; i < data.length; i++) {
        let obj = {};
        if (Object.keys(data[i]).length === 17) {
            let key = Object.values(data[i]);
            let value = Object.values(data[i + 1]);
            if (key.length === value.length) {
                obj[Object.values(data[i - 1])[0]] = Object.values(data[i - 1])[1]
                for (let x = 0; x < key.length; x++) {
                    obj[key[x]] = value[x]
                }
                record.push(obj)
                i++;
            }
        }
    }
    console.log("record is ", record)
}
getRecord()
