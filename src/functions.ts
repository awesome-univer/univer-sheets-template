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

function scrollToCell () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');

    univerAPI.executeCommand('sheet.command.scroll-to-cell', {
        range: {
            startColumn: 1,
            startRow: 99,
        }
    })
}

function scrollToTop () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');

    univerAPI.executeCommand('sheet.command.scroll-to-cell', {
        range: {
            startColumn: 0,
            startRow: 0,
        }
    })
}

function scrollToBottom () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    // @ts-expect-error
    const { rowCount } = activeSheet._worksheet.getSnapshot();
    univerAPI.executeCommand('sheet.command.scroll-to-cell', {
        range: {
            startColumn: 0,
            startRow: rowCount - 1,
        }
    })
}

function setBackground () {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    const activeWorkbook = univerAPI.getActiveWorkbook();
    if (!activeWorkbook) throw Error('activeWorkbook is not defined');
    const activeSheet = activeWorkbook.getActiveSheet();
    if (!activeSheet) throw Error('activeSheet is not defined');

    const range = activeSheet.getRange(0, 0, 1, 1);
    range?.setBackgroundColor('red');
}

function commandsListenerSwitch (el: HTMLElement) {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    if (el.tmpListener) {
        el.tmpListener.dispose();
        el.tmpListener = null;
        el.innerHTML = 'Start listening commands';
        return;
    }

    el.tmpListener = univerAPI.onCommandExecuted((command) => {
        console.log(command);
    });
    el.innerHTML = 'Stop listening commands';
    alert('Press "Ctrl + Shift + I" to open the console and do some actions in the Univer Sheets, you will see the commands in the console.');
}

function editSwitch (el: HTMLElement) {
    const univerAPI = window.univerAPI;
    if (!univerAPI) throw Error('univerAPI is not defined');

    class DisableEditError extends Error {
        constructor () {
            super('Editing is disabled');
            this.name = 'DisableEditError';
        }
    }

    if (el.tmpListener) {
        el.tmpListener.dispose();
        window.removeEventListener('error', el.errListener);
        window.removeEventListener('unhandledrejection', el.errListener);
        el.tmpListener = null;
        el.innerHTML = 'Disable edit';
        return;
    }

    el.errListener = (e: PromiseRejectionEvent | ErrorEvent) => {
        const error = e instanceof PromiseRejectionEvent ? e.reason : e.error;
        if (error instanceof DisableEditError) {
            e.preventDefault();
            console.warn('Editing is disabled');
        }
    }
    window.addEventListener('error', el.errListener);
    window.addEventListener('unhandledrejection', el.errListener);
    el.tmpListener = univerAPI.onBeforeCommandExecute(() => {
        throw new DisableEditError();
    });
    el.innerHTML = 'Enable edit';
}