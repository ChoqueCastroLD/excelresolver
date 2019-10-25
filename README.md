# excelresolver
A nodejs module for parsing excel files (xlsx, csv, xlsm, ...etc) into JavaScript Object (JSON), using promises


## Setup

Into your project folder do

> npm install excelresolver --save

And you're good to go

## Usage

Use the function *.parse* to

Extract data from a spreadsheet file
(These can be xls, xlsx, csv, xlsm, ...etc)

This will return an array of rows corresponding at the spreadsheet data

````javascript
const resolver = require('excelresolver');

const path = require('path');

// Async main function in order to use async/await syntax
async function main(){

    const myfile = path.join(__dirname, 'people.xls');
    // Excel file contains:
    //      A1       A2      A3
    //  1   id      name    lastname
    //  2   0       Luis    Choque
    //  3   1       Martha  Wayne
    //  4   2       Freddy  Mercury

    let data = await resolver.parse(myfile);
    console.table(data); // Same as console.log but with table styling

    // Will print
    // ┌─────────┬─────┬──────────┬───────────┐
    // │ (index) │ id  │   name   │ lastname  │
    // ├─────────┼─────┼──────────┼───────────┤
    // │    0    │ '0' │  'Luis'  │ 'Choque'  │
    // │    1    │ '1' │ 'Martha' │  'Wayne'  │
    // │    2    │ '2' │ 'Freddy' │ 'Mercury' │
    // └─────────┴─────┴──────────┴───────────┘

    console.log(data[0]);

    // Will print
    // {
    //     id: 0,
    //     name: 'Luis',
    //     lastname: 'Choque
    // }
}

main();
````