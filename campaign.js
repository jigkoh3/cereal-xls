var request = require('request');
var excel = require('excel4node');
const getOnlineOrders = new Promise((resolve, reject) => {
    var options = {
        'method': 'GET',
        'url': 'https://api.thamturakit.com/api/orders?channel=%E0%B8%82%E0%B8%B2%E0%B8%A2%E0%B8%AB%E0%B8%B8%E0%B9%89%E0%B8%99',
        'headers': {
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MGUxNmQ4ZDE0OWNmMzAwMTkyYzk1Y2QiLCJmaXJzdE5hbWUiOiJUaGVlcmFzYWsiLCJsYXN0TmFtZSI6IlR1YnJpdCIsInJvbGVzIjpbImFkbWluIl0sImlhdCI6MTYzNDg1NjMyNCwiZXhwIjoxNjQyMDU2MzI0fQ.yPvp-t5yJN_Vu1yYCgZOhxB4m7qRAXsZMl4jPujGuhU'
        }
    };
    request(options, function (error, response) {
        if (error) reject(error);

        const res = JSON.parse(response.body);
        const result = res.data.filter((data) => {
            return data.items[0].qty >= 100 && data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'SHAREHOLDER' && new Date(data.payments[0].transactionDateandTime) > new Date('2021-09-30');
        })
        resolve(result);

    });
});

const getAllContacts = new Promise((resolve, reject) => {
    var options = {
        'method': 'GET',
        'url': 'https://api.thamturakit.com/api/contacts',
        'headers': {
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MGUxNmQ4ZDE0OWNmMzAwMTkyYzk1Y2QiLCJmaXJzdE5hbWUiOiJUaGVlcmFzYWsiLCJsYXN0TmFtZSI6IlR1YnJpdCIsInJvbGVzIjpbImFkbWluIl0sImlhdCI6MTYzNDg1NjMyNCwiZXhwIjoxNjQyMDU2MzI0fQ.yPvp-t5yJN_Vu1yYCgZOhxB4m7qRAXsZMl4jPujGuhU'
        }
    };
    request(options, function (error, response) {
        if (error) throw new Error(error);
        let res = JSON.parse(response.body);

        resolve(res.data);
    });
})

getOnlineOrders.then((res) => {
    getAllContacts.then((contacts) => {
        // Create a new instance of a Workbook class
        var workbook = new excel.Workbook();
        var worksheet = workbook.addWorksheet('Sheet 1');


        worksheet.cell(1, 1).string('เลขที่')
        worksheet.cell(1, 2).string('วันที่ชำระเงิน')
        worksheet.cell(1, 3).string('ชื่อผู้รับ')
        worksheet.cell(1, 4).string('เบอร์โทร')
        worksheet.cell(1, 5).string('Line name')
        worksheet.cell(1, 6).string('ที่อยู่จักส่ง')
        worksheet.cell(1, 7).string('จำนวน(หุ้น)')

        let i = 2;
        res.forEach(order => {

            // console.log(order)
            const xx = contacts.find(ct => ct._id === order.customerId);
            worksheet.cell(i, 1).string(order.id)
            worksheet.cell(i, 2).string(`${order.payments[0].transactionDateandTime}`)
            worksheet.cell(i, 3).string(` ${xx.title} ${xx.firstName} ${xx.lastName}`)
            worksheet.cell(i, 4).string(xx.tel)
            worksheet.cell(i, 5).string(xx.lineName)
            worksheet.cell(i, 6).string(`${xx.addr01} ${xx.street} ${xx.subDistrict} ${xx.district} ${xx.province} ${xx.zip}`)
            worksheet.cell(i, 7).number(order.items[0].qty)

            i++;
        });

        workbook.write(`รายการจัดส่งาวหอมมะลิธรรมชาติแท้ใหม่ต้นฤดูที่ซื้อหุ้นตั้งแต่100หุ้นขึ้นไป.xlsx`);
    });
})