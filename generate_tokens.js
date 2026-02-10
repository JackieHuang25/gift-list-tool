const fs = require('fs');
const csv = require('csv-parser');
const crypto = require('crypto');

const results = [];

const headerMap = {
  '*Order Number': 'Order Number',
  'Backer No.': 'Double Check No',
  '*Reward Name': 'Reward Name',
  '*Item Quantity': 'Item Quantity',
  '*Full Name': 'Full Name',
  'Phone Number': 'Phone Number',
  '*Country/Region Code': 'Country/Region Code',
  'State/Province/Region': 'State/Province/Region',
  '*City': 'City',
  '*Address1': 'Address1',
  'Address2': 'Address2',
  'Zip Code': 'Zip Code',
  'Email': 'Email',
  'Total Payment': 'Total Payment',
  'Match Status': 'Match Status',
  'Changed Fields': 'Changed Fields',
  'Changed Values': 'Changed Values',
  'Submitted At': 'Submitted At'
};

function normalizeRow(row) {
  const normalized = {};
  Object.keys(row).forEach((key) => {
    const mappedKey = headerMap[key] || key;
    normalized[mappedKey] = row[key];
  });

  return normalized;
}

// 读取 CSV
fs.createReadStream('gift_list.csv')
  .pipe(csv())
  .on('data', (data) => results.push(normalizeRow(data)))
  .on('end', () => {
    // 给每条记录生成 token + expire
    const giftListWithToken = results.map(item => {
      const token = crypto.randomBytes(16).toString('hex'); // 32 位随机 token
      const expire = Date.now() + 7 * 24 * 60 * 60 * 1000; // 7 天有效期
      return {...item, token, expire};
    });

    // 输出 JSON 文件
    fs.writeFileSync('gift_list_token.json', JSON.stringify(giftListWithToken, null, 2));
    console.log('gift_list_token.json generated with tokens!');
  });
