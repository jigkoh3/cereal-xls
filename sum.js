const request = require('request');
const excel = require('excel4node');
const header = {
    'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MGUxNmQ4ZDE0OWNmMzAwMTkyYzk1Y2QiLCJmaXJzdE5hbWUiOiJUaGVlcmFzYWsiLCJsYXN0TmFtZSI6IlR1YnJpdCIsInJvbGVzIjpbImFkbWluIl0sImlhdCI6MTY0MjE3MDg5MSwiZXhwIjoxNjQ5MzcwODkxfQ.ZEGH9NprHgqx6ivms1DWXVIuVTNmkTQY6YDSptXLXVU'
};

const getAllShareHolders = new Promise((resolve, reject) => {
    var options = {
        'method': 'GET',
        'url': 'https://api.thamturakit.com/api/shareholdercerts',
        'headers': header
    };
    request(options, function (error, response) {
        if (error) throw new Error(error);

        let res = JSON.parse(response.body);
        resolve(res.data);
    });
})

const getAllContacts = new Promise((resolve, reject) => {
    var options = {
        'method': 'GET',
        'url': 'https://api.thamturakit.com/api/contacts',
        'headers': header
    };
    request(options, function (error, response) {
        if (error) throw new Error(error);
        let res = JSON.parse(response.body);

        resolve(res.data);
    });
})

getAllShareHolders.then(res => {
    var workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet('Sheet 1');


    worksheet.cell(1, 1).string('เลขที่ใบหุ้น')
    worksheet.cell(1, 2).string('วันที่')
    worksheet.cell(1, 3).string('เลขที่ผู้ถือหุ้น')
    worksheet.cell(1, 4).string('ชื่อผู้ถือหุ้น')
    worksheet.cell(1, 5).string('จำนวนหุ้นที่ซื้อ')
    worksheet.cell(1, 6).string('มูลค่าหุ้นที่ซื้อ')


    let i = 2;
    res.forEach(shareholder => {
        // console.log(shareholder);
        try {
            
            worksheet.cell(i, 1).string(shareholder.no)
            worksheet.cell(i, 2).string(shareholder.created)
            worksheet.cell(i, 3).string(shareholder.shareHolderNo)
            worksheet.cell(i, 4).string(shareholder.name)
            worksheet.cell(i, 5).number(shareholder.qty)
            worksheet.cell(i, 6).number(shareholder.amount)
        } catch (error) {
            console.log(error)
            console.log(shareholder);
        }


        i++;
    });
    workbook.write(`รายงานซื้อหุ้นตั่งแต่เริ่มกิจการ.xlsx`);
})