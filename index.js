var request = require('request');
var excel = require('excel4node');
const getOnlineOrders = new Promise((resolve, reject) => {
    var options = {
        'method': 'GET',
        'url': 'https://api.thamturakit.com/api/orders?channel=%E0%B8%82%E0%B8%B2%E0%B8%A2%E0%B8%AD%E0%B8%AD%E0%B8%99%E0%B9%84%E0%B8%A5%E0%B8%99%E0%B9%8C',
        'headers': {
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MGUxNmQ4ZDE0OWNmMzAwMTkyYzk1Y2QiLCJmaXJzdE5hbWUiOiJUaGVlcmFzYWsiLCJsYXN0TmFtZSI6IlR1YnJpdCIsInJvbGVzIjpbImFkbWluIl0sImlhdCI6MTYzNDg1NjMyNCwiZXhwIjoxNjQyMDU2MzI0fQ.yPvp-t5yJN_Vu1yYCgZOhxB4m7qRAXsZMl4jPujGuhU'
        }
    };
    request(options, function (error, response) {
        if (error) reject(error);

        const res = JSON.parse(response.body);
        const result = res.data.filter((data) => {
            return data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'CEREALORDER' && data.payments[0].transactionDateandTime.startsWith('2021-12-01')
            // return data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'CEREALORDER' && ['2021111603012556','2021111603004624','2021111602594698','2021111601254139'].includes(data.id)
        })
        resolve(result);

    });
});

getOnlineOrders.then((res) => {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet('Sheet 1');

   
    worksheet.cell(1, 1).string('เลขที่')
    worksheet.cell(1, 2).string('วันที่ชำระเงิน')
    worksheet.cell(1, 3).string('ชื่อผู้รับ')
    worksheet.cell(1, 4).string('เบอร์โทร')
    worksheet.cell(1, 5).string('Line name')
    worksheet.cell(1, 6).string('ที่อยู่จักส่ง')
    worksheet.cell(1, 7).string('จำนวน(กล่อง)')
   

    let i = 2;
    res.forEach(order => {
        
        // console.log(order)

        worksheet.cell(i, 1).string(order.id)
        worksheet.cell(i, 2).string(`${order.payments[0].transactionDateandTime}`)
        worksheet.cell(i, 3).string(` ${order.shipments[0].title} ${order.shipments[0].firstName} ${order.shipments[0].lastName}`)
        worksheet.cell(i, 4).string(order.shipments[0].tel)
        worksheet.cell(i, 5).string(order.shipments[0].lineName)
        worksheet.cell(i, 6).string(`${order.shipments[0].addr01} ${order.shipments[0].street} ${order.shipments[0].subDistrict} ${order.shipments[0].district} ${order.shipments[0].province} ${order.shipments[0].zip}` )
        worksheet.cell(i, 7).number(order.items[0].qty)
        
        i++;
    });

    workbook.write('รายการจัดส่งซีเรียสซีเรียล-20211201.xlsx');
})




