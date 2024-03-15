function setValue () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const value = 'Hello, World!';

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    const range = activeSheet.getRange(0, 0);
    if (!range) throw Error('range is not defined');

    /** 
     * @see https://univer.ai/api/facade/classes/FRange.html#setValue
     */
    range.setValue(value);
}

function setValues () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const values = [
        ['Hello', 'World!'],
        ['Hello', 'Univer!'],
        ['Hello', 'Sheets!']
    ]

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    const range = activeSheet.getRange(0, 0, values.length, values[0].length);
    if (!range) throw Error('range is not defined');

    /** 
     * @see https://univer.ai/api/facade/classes/FRange.html#setValues
     */
    range.setValues(values);
}

function getValue () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const values = [
        ['Hello', 'World!'],
        ['Hello', 'Univer!'],
        ['Hello', 'Sheets!']
    ]

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    const range = activeSheet.getRange(0, 0, values.length, values[0].length);
    if (!range) throw Error('range is not defined');

    /** 
     * @see https://univer.ai/api/facade/classes/FRange.html#getValue
     */
    alert(JSON.stringify(range.getValue(), null, 2));
    console.log(JSON.stringify(range.getValue(), null, 2));
}

function getValues () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const values = [
        ['Hello', 'World!'],
        ['Hello', 'Univer!'],
        ['Hello', 'Sheets!']
    ]

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    const range = activeSheet.getRange(0, 0, values.length, values[0].length);
    if (!range) throw Error('range is not defined');

    // TODO: add facade API
    const data: (string | undefined)[][] = [];
    range.forEach((row, col, cell) => {
        data[row] = data[row] || [];
        data[row][col] = cell.v?.toString();
    });

    alert(JSON.stringify(data, null, 2));
    console.log(JSON.stringify(data, null, 2));
}


function getWorkbookData () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');


    alert(JSON.stringify(activeWorkbook.getSnapshot(), null, 2));
    console.log(JSON.stringify(activeWorkbook.getSnapshot(), null, 2));
}

function getSheetData () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');

    const snapshot = activeWorkbook.getSnapshot();
    const sheet1 = Object.values(snapshot.sheets).find((sheet) => {
        return sheet.name === 'Sheet1';
    });

    if (!sheet1) throw Error('sheet1 is not defined');
    alert(JSON.stringify(sheet1, null, 2));
    console.log(JSON.stringify(sheet1, null, 2));
}

function createSheet () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');

    const sheet = activeWorkbook.create("Sheet2", 10, 10);

    if (!sheet) throw Error('sheet is not defined');
    alert('Sheet created');
}