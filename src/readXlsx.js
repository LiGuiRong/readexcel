const xlsx = require('node-xlsx').default;
const fs = require('fs');

const importData = [];

const dataMapKey = {
    "姓名": null,
    "年龄": null,
    "性别": null,
    "职位": null,
}

const workSheetsFromBuffer = xlsx.parse(fs.readFileSync('./test.xlsx'));
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


const exportData = [
    {
        name: 'sheet1',
        data: [
            ['姓名', '年龄', '性别'], // 表格标题
            ['李贵荣', '25', '男'], // 第二行开始是表格数据
        ]
    },
    {
        name: 'sheet2',
        data: [
            ['姓名', '年龄', '性别'], // 表格标题
            ['李贵荣', '25', '男'], // 第二行开始是表格数据
        ]
    }
];

const buffer = xlsx.build(exportData);
fs.writeFile('./src/output.xlsx', buffer, function(err) {
    if (err) {
        throw err;
    }
    console.log('successfully!') 
})
