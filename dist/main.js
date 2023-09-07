/**
/* main.ts
/* - XLSXAggregator -
**/
"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// modules
const electron_1 = __importDefault(require("electron")); // electron
const xlsx_1 = __importDefault(require("xlsx")); // xlsx
const path_1 = __importDefault(require("path")); // path
const fs = __importStar(require("fs")); // fs
// functions
const functions = {
    // remove unnecessary symbols
    removeSymbol: (str) => {
        return new Promise((resolve, reject) => {
            try {
                // tmp text
                let tmpstr = str;
                // arguments
                const args = ["\\", "\/", "?", "\*", "\[", "\]"];
                // loop for arguments
                args.forEach(arg => {
                    // contains arg
                    if (tmpstr.indexOf(arg) != -1) {
                        // remove them
                        tmpstr.replace(arg, '');
                    }
                });
                // resolved
                resolve(tmpstr);
            }
            catch (e) {
                // reject
                reject(e.toString());
            }
        });
    },
    // get formatted date
    getFormattedDate: (date, format) => {
        // symbols
        const symbol = {
            M: date.getMonth() + 1,
            d: date.getDate(),
            h: date.getHours(),
            m: date.getMinutes(),
            s: date.getSeconds(),
        };
        // fromatted strings
        const formatted = format.replace(/(M+|d+|h+|m+|s+)/g, (v) => ((v.length > 1 ? "0" : "") + symbol[v.slice(-1)]).slice(-2));
        // return symbols
        return formatted.replace(/(y+)/g, (v) => date.getFullYear().toString().slice(-v.length));
    }
};
// now date
const now = new Date;
// output string
const OUTPUT_STR = functions.getFormattedDate(now, "yyyymmdd");
// target folder
const OUT_PATH = 'output';
// target file path
const OUT_FILEPATH = `${OUT_PATH}/${OUTPUT_STR}.csv`;
// list_office title
const OUT_TITLE = [
    ["送信者タイプ", "送信者名", "送信日", "送信時刻", "内容"]
];
// main
electron_1.default.app.on('ready', () => __awaiter(void 0, void 0, void 0, function* () {
    // create workbook
    const outWb = xlsx_1.default.utils.book_new();
    // empty sheet
    const outWs = xlsx_1.default.utils.json_to_sheet([]);
    // select target dir
    const dir = yield electron_1.default.dialog.showOpenDialog({
        // select directory
        properties: ['openDirectory'],
    });
    // make folder
    if (!fs.existsSync(OUT_PATH)) {
        fs.mkdirSync(OUT_PATH);
    }
    // result object
    const promiseObj = yield fileLister(dir.filePaths[0]);
    // append dummy sheet
    xlsx_1.default.utils.book_append_sheet(outWb, outWs, OUT_PATH);
    // add sheet
    xlsx_1.default.utils.sheet_add_aoa(outWb, OUT_TITLE, { origin: { r: 0, c: 0 } });
    // get worksheet      
    const outOutWs = outWb.Sheets[outWb.SheetNames[0]];
    // append data
    xlsx_1.default.utils.sheet_add_aoa(outOutWs, promiseObj, { origin: { r: 1, c: 0 } });
    // append sheet
    xlsx_1.default.writeFile(outWb, OUT_FILEPATH);
    // show message
    electron_1.default.dialog.showMessageBox({
        type: 'info',
        message: '完了',
    });
}));
// list filenames
const fileLister = (dirpath) => {
    return new Promise((resolve, reject) => {
        // read all files
        fs.readdir(dirpath, { withFileTypes: true }, (err, dirents) => __awaiter(void 0, void 0, void 0, function* () {
            // result array
            let resultArray = [];
            // error
            if (err)
                reject(err);
            // loop inside directory
            for (const dirent of dirents) {
                // target path
                const fp = path_1.default.join(dirpath, dirent.name);
                // is directory
                if (dirent.isDirectory()) {
                    // list file names
                    fileLister(fp);
                    // is file
                }
                else {
                    // get workbook
                    const originWb = xlsx_1.default.readFile(fp);
                    // get sheetlist
                    const originWsList = originWb.SheetNames;
                    // loop sheetname
                    for (const ls of originWsList) {
                        // sheetname
                        const sheetname = yield functions.removeSymbol(ls);
                        // get sheet1
                        const originWs = originWb.Sheets[sheetname];
                        // convert to json
                        const originOutWsJson = xlsx_1.default.utils.sheet_to_json(originWs, {
                            defval: "", // blank ok
                        });
                        // write log
                        for (let j = 3; j < originOutWsJson.length; j++) {
                            // all data
                            Object.defineProperty(originOutWsJson[j], 'file', {
                                value: dirent.name,
                                writable: false,
                                enumerable: true,
                                configurable: false
                            });
                        }
                        // write log
                        for (let i = 3; i < originOutWsJson.length; i++) {
                            // all data
                            resultArray.push(Object.keys(originOutWsJson[i]).map((key) => {
                                return String(originOutWsJson[i][key]).replace(/\r?\n/g, '');
                            }));
                        }
                    }
                }
            }
            console.log(resultArray);
            // resolved
            resolve(resultArray);
        }));
    });
};
