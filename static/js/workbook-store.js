const WORKBOOK_DB_NAME = "excel-analyzer-storage";
const WORKBOOK_STORE_NAME = "workbooks";
const WORKBOOK_KEY = "last-workbook";

function abrirAlmacenWorkbook() {
    return new Promise(function(resolve, reject) {
        const request = indexedDB.open(WORKBOOK_DB_NAME, 1);
        request.onupgradeneeded = function() {
            request.result.createObjectStore(WORKBOOK_STORE_NAME);
        };
        request.onsuccess = function() { resolve(request.result); };
        request.onerror = function() { reject(request.error); };
    });
}

async function guardarWorkbookLocal(buffer, name) {
    const database = await abrirAlmacenWorkbook();
    return new Promise(function(resolve, reject) {
        const transaction = database.transaction(WORKBOOK_STORE_NAME, "readwrite");
        transaction.objectStore(WORKBOOK_STORE_NAME).put({ buffer, name }, WORKBOOK_KEY);
        transaction.oncomplete = function() {
            database.close();
            resolve();
        };
        transaction.onerror = function() {
            database.close();
            reject(transaction.error);
        };
    });
}

async function leerWorkbookLocal() {
    const database = await abrirAlmacenWorkbook();
    return new Promise(function(resolve, reject) {
        const transaction = database.transaction(WORKBOOK_STORE_NAME, "readonly");
        const request = transaction.objectStore(WORKBOOK_STORE_NAME).get(WORKBOOK_KEY);
        request.onsuccess = function() {
            database.close();
            resolve(request.result || null);
        };
        request.onerror = function() {
            database.close();
            reject(request.error);
        };
    });
}
