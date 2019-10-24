const fs = require('fs');
const readExcel = require('read-excel-file');
const excelJson = require("sheet2json");

function promiseExcelJson(excel) {
    return new Promise((resolve, reject) => {
        try {
            excelJson(excel, (res) => {
                resolve(res);
            });
        } catch (error) {
            reject(error);
        }
    })
}

Promise.properRace = function (...promises) {
    if (promises.length < 1) {
        return Promise.reject('Can\'t start a race without promises!');
    }
    let indexPromises = promises.map((p, index) => p.catch(() => {
        throw index;
    }));

    return Promise.race(indexPromises).catch(index => {
        let p = promises.splice(index, 1)[0];
        p.catch(e => console.log('A car has crashed, don\'t interrupt the race:', e));
        return Promise.properRace(promises);
    });
};

async function parseExcel(path) {
    try {
        let ret = await promiseExcelJson(path);
        return ret;
    } catch (error) {
        let file = fs.readFileSync(path);
        let ret = await readExcel(file);
        return ret;
    }
}

async function postParseExcel(path){
    let data = await parseExcel(path);
    
    if(Reflect.has(data, 'json'))
        data = data.json;

    let res = [];

    for (let d of data) {
        let temp = {};

        if(Reflect.has(d, 'json'))
            d = d.json;

        for (const key in d) {
            temp[key.toString().toLowerCase().split('.').join('').split(' ').join('_').split('"').join('').split('"').join('')] = d[key];
        }

        res.push(temp);
    }

    if(Array.isArray(res[0].columns)){
        let temp = [];
        for (const index in res[0])
            if(!Array.isArray(res[0][index]))
                temp.push(res[0][index]);
        res = temp;
    }

    return res;
}

module.exports = {
    /**
     * 
     * Parse your excel file into a Javascript Object
     * 
     * @param {String} path The absolute path of the file to parse
     */
    parse: async (path) => await postParseExcel(path)
};