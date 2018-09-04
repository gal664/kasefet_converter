const xlsx = require('node-xlsx');
const fs = require('fs');

let curDate = new Date().toISOString().split(".")[0].split(":").join("-").split("T").join("-");
let obj = xlsx.parse('../converter_sheets.xlsx');
let semelMosad = obj[0].data[2][1]
let shnatLimudim = obj[0].data[3][1]

for (let sn = 2; sn < obj.length; sn++) {
      let writeStr = "";
      let rows = [];
      let sheet = obj[sn];
      let fileName = `${sheet.name}_${shnatLimudim}_${semelMosad}_${curDate}_from_moe.csv`

      for (let j = 0; j < sheet['data'].length; j++) {
            rows.push(sheet['data'][j]);
      }

      for (let i = 0; i < rows.length; i++) {
            if (rows[i][0] === "" || rows[i][0] === undefined) {
                  rows.splice(i, rows.length);
                  break;
            }
            writeStr += rows[i].join(",") + "\n";
      }

      fs.writeFile(`../files/${fileName}`, writeStr, () => {
            console.log(`created ${fileName}`);
      });
}