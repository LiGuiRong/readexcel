const express = require('express');
const app = express();
const fs = require("fs");
const bodyParser = require('body-parser');
const multer  = require('multer');
const xlsx = require('node-xlsx');
const path = require('path');
const mime = require('mime');

const upload = multer({ dest: './uploads/' })

app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: false }));

app.get('/index.html', function (req, res) {
   res.sendFile( __dirname + "/" + "index.html" );
})

function handleMimeType(mimetype) {
    if (mimetype.indexOf('javascript') > -1) return 'js';
    return mime.getExtension(mimetype);
}

//  注意upload.single('uploadFile')的参数uploadFile是input的name属性，不然会报错
app.post('/file_upload', upload.single('uploadFile'), function (req, res) {
    res.writeHead(200, {'Content-Type': 'text/plain; charset=utf-8'}); // 数据编码格式
    console.log('req.file', req.file);  // 上传的文件信息
    const absolutePath = req.file.path + '.' + handleMimeType(req.file.mimetype);
    console.log('mimetype', absolutePath);
    fs.rename(req.file.path, absolutePath, function(err) {
        if (err) {
            throw err;
        }
    })
    const response = {
        message:'File uploaded successfully', 
        filename: req.file.filename,
        filePath: path.resolve(absolutePath)
    };
    res.end( JSON.stringify( response ) );
})

const dataMapKey = {
    "姓名": null,
    "年龄": null,
    "性别": null,
    "职位": null,
}

//  注意upload.single('uploadFile')的参数uploadFile是input的name属性，不然会报错
app.post('/read_xlsx', upload.single('uploadFile'), function (req, res) {
    res.writeHead(200, {'Content-Type': 'text/plain; charset=utf-8'}); // 数据编码格式
    console.log('req.file', req.file);  // 上传的文件信息
    // 获取到的信息
    // req.file = { fieldname: 'uploadFile',
    //     originalname: 'test.xlsx',
    //     encoding: '7bit',
    //     mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    //     destination: './uploads/',
    //     filename: '5d18655c24ee6cca753b981fd8d9ad24',
    //     path: 'uploads\\5d18655c24ee6cca753b981fd8d9ad24',
    //     size: 10876 
    // }
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(req.file.path));
    const importData = [];
    workSheetsFromBuffer.forEach(item => {
        item.data[0].forEach((subItem, index) => {
            dataMapKey[subItem] = index;
        })
        console.log('表格标题', dataMapKey);
        item.data.forEach((sub, index) => {
            if (index > 0) {    // 从第二行开始读取，第一行为标题
                importData.push({
                    name: sub[dataMapKey['姓名']] || '',
                    age: sub[dataMapKey['年龄']] || '',
                    sex: sub[dataMapKey['性别']] || '',
                    position: sub[dataMapKey['职位']] || '',
                })
            }
        })
    })

    console.log(importData);
    // 此出可执行数据库写入功能

    // 将文件重命名
    const absolutePath = req.file.path + '.' + handleMimeType(req.file.mimetype);
    fs.rename(req.file.path, absolutePath, function(err) {
        if (err) {
            throw err;
        }
    })
    const response = {
        message:'File uploaded successfully', 
        filename: req.file.filename,
        filePath: path.resolve(absolutePath),
        importData: importData
    };
    res.end( JSON.stringify( response ) );
})

var server = app.listen(8081, function () {
    var host = server.address().address
    var port = server.address().port
})