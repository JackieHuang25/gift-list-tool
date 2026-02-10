// pages/api/getAddress.js
import fs from 'fs';
import path from 'path';

export default function handler(req,res){
  const { token } = req.query;
  const filePath = path.join(process.cwd(),'gift_list_token.json');
  const giftList = JSON.parse(fs.readFileSync(filePath));

  const record = giftList.find(r=>r.token===token);
  if(!record) return res.status(404).json({error:'Invalid token'});
  if(record.expire < Date.now()) return res.status(403).json({error:'Token expired'});

  res.status(200).json(record);
}
