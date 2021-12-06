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
            // return data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'CEREALORDER' && data.payments[0].transactionDateandTime.startsWith('2021-11-20')
            // return data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'CEREALORDER' && ['2021111603012556','2021111603004624','2021111602594698','2021111601254139'].includes(data.id)
            return data.payments.length > 0 && data.payments[0].billPaymentRef2 === 'SHAREHOLDER' && new Date(data.payments[0].transactionDateandTime) > new Date('2021-06-30');
        })
        resolve(result);

    });
});

getOnlineOrders.then((res) => {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet('Sheet 1');


    worksheet.cell(1, 1).string('วันที่')
    worksheet.cell(1, 2).string('จำนวนเงินรวม')





    var result = [];
    res.reduce(function (res, value) {
        if (!res[value.payments[0].transactionDateandTime.substring(0, 10)]) {
            res[value.payments[0].transactionDateandTime.substring(0, 10)] = { trnsDate: value.payments[0].transactionDateandTime.substring(0, 10), amount: value.netAmount }

            result.push(res[value.payments[0].transactionDateandTime.substring(0, 10)])
        }

        res[value.payments[0].transactionDateandTime.substring(0, 10)].amount += value.netAmount;
        return res;
    }, {});

    console.log(result);

    let i = 2;
    result.forEach(order => {

        console.log(order)

        worksheet.cell(i, 1).string(`${order.trnsDate}`)
        worksheet.cell(i, 2).string(`${order.amount}`)

        i++;
    });

    workbook.write(`สรุปยอดรวมซื้อหุ้นรายวัน.xlsx`);

    // console.log(total);
})




