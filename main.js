/**
/* main.js
/* - XLSXAggregator -
**/

"use strict";

// modules
const { app, dialog } = require('electron'); // electron
const xlsx = require('xlsx'); // xlsx
const path = require('path'); // path
const fs = require('fs'); // file operator
const xutil = xlsx.utils; // xlsx utility

// string
const OFFICE_STR = 'list_office';
const SHOP_STR = 'list_shop';

// output path
const OUT_PATH = './output';
const OUT_OFFICE_FILEPATH = `${OUT_PATH}/${OFFICE_STR}.xlsx`;
const OUT_SHOP_FILEPATH = `${OUT_PATH}/${SHOP_STR}.xlsx`;

// list_office title
const OFFICE_TITLE = [
    ["*****", "*****", "", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****"],
];
// list_shop title
const SHOP_TITLE = [
    ["*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****", "*****"],
];
        
// main
app.on('ready', async() => {
    // create workbook
    const officeWb = xutil.book_new();
    const shopWb = xutil.book_new();
    // empty sheet
    const officeWs = xutil.json_to_sheet([]);
    const shopWs = xutil.json_to_sheet([]);

    // select target dir
    const dir = await dialog.showOpenDialog(null, {
        // select directory
        properties: ['openDirectory'], 
    });

    // make folder
    if(!fs.existsSync(OUT_PATH)) {
        fs.mkdirSync(OUT_PATH);
    }

    // result object
    const promiseObj = await fileLister(dir.filePaths[0]);
    // final array
    const lastOfficeArray = promiseObj.office.map(JSON.stringify).reverse().filter((e, i, a) => a.indexOf(e, i + 1) === -1).reverse().map(JSON.parse);
    const lastShopArray = promiseObj.shop.map(JSON.stringify).reverse().filter((e, i, a) => a.indexOf(e, i + 1) === -1).reverse().map(JSON.parse);
    
    // append dummy sheet
    xutil.book_append_sheet(officeWb, officeWs, OFFICE_STR);
    xutil.book_append_sheet(shopWb, shopWs, SHOP_STR);
    // add sheet
    xutil.sheet_add_aoa(officeWs, OFFICE_TITLE,{ origin: { r: 0, c: 0 } });
    xutil.sheet_add_aoa(shopWs, SHOP_TITLE,{ origin: { r: 0, c: 0 } });

    // get worksheet      
    const outOfficeWs = officeWb.Sheets[officeWb.SheetNames[0]];
    const outShopWs = shopWb.Sheets[shopWb.SheetNames[0]];

    // office
    // append data
    xutil.sheet_add_aoa(outOfficeWs, lastOfficeArray,{ origin: { r: 1, c: 0 } });
    // append sheet
    xlsx.writeFile(officeWb, OUT_OFFICE_FILEPATH);

    // shop
    // append data
    xutil.sheet_add_aoa(outShopWs, lastShopArray,{ origin: { r: 1, c: 0 } });
    // append sheet
    xlsx.writeFile(shopWb, OUT_SHOP_FILEPATH);
    
    // show message
    dialog.showMessageBox(null, {
        type: 'info',
        message: 'completed',
    });
});

// list filenames
const fileLister = dirpath => {
    return new Promise((resolve, reject) => {
        // read all files
        fs.readdir(dirpath, { withFileTypes: true }, (err, dirents) => {
            // result array
            let resultOfficeArray = [];
            let resultShopArray = [];

            // error
            if(err) reject(err);
            
            // loop inside directory
            for(const dirent of dirents) { 
                // target path
                const fp = path.join(dirpath, dirent.name);

                // is directory
                if(dirent.isDirectory()) {
                    // list file names
                    fileLister(fp);

                // is file
                } else {
                    // get workbook
                    const originWb = xlsx.readFile(fp);
                    // get sheetlist
                    const originWsList = originWb.SheetNames;
                
                    // loop sheetname
                    for(const ls of originWsList) {
                        // get sheet1
                        const originWs = originWb.Sheets[ls];
                        // if list_office
                        if(ls.includes(OFFICE_STR)) {
                            // json
                            const originOfficeWsJson = xutil.sheet_to_json(originWs, {
                                defval: "", // blank ok
                            });
                            // write log
                            for(let i = 0; i < originOfficeWsJson.length; i++ ) {
                                // all data
                                resultOfficeArray.push(Object.keys(originOfficeWsJson[i]).map(key => { return originOfficeWsJson[i][key] }));
                            }
                            
                        // if list_shop
                        } else if(ls.includes(SHOP_STR)) {
                            // json
                            const originShopWsJson = xutil.sheet_to_json(originWs, {
                                defval: "", // blank ok
                            });
                            // write log
                            for(let j = 0; j < originShopWsJson.length; j++ ) {
                                // all data
                                resultShopArray.push(Object.keys(originShopWsJson[j]).map(key => { return originShopWsJson[j][key] }));
                            }
                        }
                    }
                }
            } 
            // resolved
            resolve({
                office: resultOfficeArray, 
                shop: resultShopArray,
            });
        }); 
    }); 
}
