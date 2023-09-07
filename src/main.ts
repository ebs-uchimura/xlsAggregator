/**
/* main.ts
/* - XLSXAggregator -
**/

"use strict";

// modules
import electron from 'electron'; // electron
import xlsx from 'xlsx'; // xlsx
import path from 'path'; // path
import * as fs from 'fs'; // fs

// functions
const functions = {
    // remove unnecessary symbols
    removeSymbol : (str: string): Promise<string> => {
        return new Promise((resolve, reject) => {
            try {
                // tmp text
                let tmpstr = str;
                // arguments
                const args: string[] = ["\\", "\/", "?", "\*", "\[", "\]"];
                // loop for arguments
                args.forEach(arg => {
                    // contains arg
                    if ( tmpstr.indexOf(arg) != -1) {
                        // remove them
                        tmpstr.replace(arg, '');
                    }
                });
                // resolved
                resolve(tmpstr);

            } catch(e: any) {
                // reject
                reject(e.toString());
            }
        });
    },

    // get formatted date
    getFormattedDate : (date: Date, format: string) => {
        // symbols
        const symbol = {
            M: date.getMonth() + 1,
            d: date.getDate(),
            h: date.getHours(),
            m: date.getMinutes(),
            s: date.getSeconds(),
        };

        // fromatted strings
        const formatted = format.replace(/(M+|d+|h+|m+|s+)/g, (v) =>
            ((v.length > 1 ? "0" : "") + symbol[v.slice(-1) as keyof typeof symbol]).slice(-2)
        )

        // return symbols
        return formatted.replace(/(y+)/g, (v) =>
            date.getFullYear().toString().slice(-v.length)
    )}
}

// now date
const now:Date = new Date;
// output string
const OUTPUT_STR: string = functions.getFormattedDate(now, "yyyymmdd");
// target folder
const OUT_PATH = 'output';
// target file path
const OUT_FILEPATH = `${OUT_PATH}/${OUTPUT_STR}.csv`;

// list_office title
const OUT_TITLE = [
    ["送信者タイプ", "送信者名", "送信日", "送信時刻", "内容"]
];
        
// main
electron.app.on('ready', async() => {
    // create workbook
    const outWb = xlsx.utils.book_new();
    // empty sheet
    const outWs = xlsx.utils.json_to_sheet([]);
    // select target dir
    const dir = await electron.dialog.showOpenDialog({
        // select directory
        properties: ['openDirectory'], 
    });
    // make folder
    if(!fs.existsSync(OUT_PATH)) {
        fs.mkdirSync(OUT_PATH);
    }
    // result object
    const promiseObj: any = await fileLister(dir.filePaths[0]);
    // append dummy sheet
    xlsx.utils.book_append_sheet(outWb, outWs, OUT_PATH);
    // add sheet
    xlsx.utils.sheet_add_aoa(outWb, OUT_TITLE, { origin: { r: 0, c: 0 } });
    // get worksheet      
    const outOutWs = outWb.Sheets[outWb.SheetNames[0]];
    // append data
    xlsx.utils.sheet_add_aoa(outOutWs, promiseObj, { origin: { r: 1, c: 0 } });
    // append sheet
    xlsx.writeFile(outWb, OUT_FILEPATH);
    
    // show message
    electron.dialog.showMessageBox({
        type: 'info',
        message: '完了',
    });
});

// list filenames
const fileLister = (dirpath: string):Promise<string[][]> => {
    return new Promise((resolve, reject) => {
        // read all files
        fs.readdir(dirpath, { withFileTypes: true }, async(err, dirents) => {
            // result array
            let resultArray: string[][] = [];
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
                        // sheetname
                        const sheetname: string = await functions.removeSymbol(ls);
                        // get sheet1
                        const originWs = originWb.Sheets[sheetname];
                        // convert to json
                        const originOutWsJson: string[] = xlsx.utils.sheet_to_json(originWs, {
                            defval: "", // blank ok
                        });
                        // write log
                        for(let j: number = 3; j < originOutWsJson.length; j++ ) {
                            // all data
                            Object.defineProperty(originOutWsJson[j], 'file', {
                                value: dirent.name,
                                writable: false, // 書き込み禁止
                                enumerable: true,
                                configurable: false
                            });
                        }
                        // write log
                        for(let i: number = 3; i < originOutWsJson.length; i++ ) {
                            // all data
                            resultArray.push(Object.keys(originOutWsJson[i]).map((key: any) => {
                                return String(originOutWsJson[i][key]).replace(/\r?\n/g, '');
                            }));
                        }
                    }
                }
            }
            console.log(resultArray);
            // resolved
            resolve(resultArray);
        }); 
    }); 
}
