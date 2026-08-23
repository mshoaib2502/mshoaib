const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { initDB, getProducts, addProduct, createInvoice, getInvoices } = require('./database');
const { getCustomers, addCustomer, searchCustomers } = require('./database');

let win;
app.whenReady().then(() => {
  initDB();
  win = new BrowserWindow({
    width: 1200, height: 800,
    webPreferences: { 
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  win.loadURL('http://localhost:5173'); // vite dev
  // win.loadFile('frontend/dist/index.html'); // for build
});

ipcMain.handle('get-products', () => getProducts());
ipcMain.handle('add-product', (e, p) => addProduct(p));
ipcMain.handle('create-invoice', (e, inv) => createInvoice(inv));
ipcMain.handle('get-invoices', () => getInvoices());
ipcMain.handle('delete-product', (e, id) => deleteProduct(id));
ipcMain.handle('get-customers', () => getCustomers());
ipcMain.handle('add-customer', (e, c) => addCustomer(c));
ipcMain.handle('search-customers', (e, q) => searchCustomers(q));