import React from 'react';
import ReactDOM from 'react-dom';
import { app, BrowserWindow, ipcMain } from 'electron';
import path from 'path';
import Datastore from 'nedb';

const db = {
  invoices: new Datastore({ filename: 'invoices.db', autoload: true }),
  items: new Datastore({ filename: 'items.db', autoload: true }),
  customers: new Datastore({ filename: 'customers.db', autoload: true }),
};

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  win.loadFile('views/index.html');
}

app.whenReady().then(createWindow);

ipcMain.handle('save-invoice', (_, payload) => {
  db.invoices.insert(payload, (err, doc) => {
    if (err) console.log(err);
    else console.log(doc);
  });
});

ipcMain.handle('list-invoices', () => {
  return new Promise((resolve, reject) => {
    db.invoices.find({}, (err, docs) => {
      if (err) reject(err);
      else resolve(docs);
    });
  });
});

ipcMain.handle('add-item', (_, { name, price }) => {
  db.items.insert({ name, price }, (err, doc) => {
    if (err) console.log(err);
    else console.log(doc);
  });
});

ipcMain.handle('list-items', () => {
  return new Promise((resolve, reject) => {
    db.items.find({}, (err, docs) => {
      if (err) reject(err);
      else resolve(docs);
    });
  });
});

ipcMain.handle('get-customer', (_, mobile) => {
  return new Promise((resolve, reject) => {
    db.customers.findOne({ mobile }, (err, doc) => {
      if (err) reject(err);
      else resolve(doc);
    });
  });
});

ipcMain.handle('save-customer', (_, { name, mobile, address }) => {
  db.customers.insert({ name, mobile, address }, (err, doc) => {
    if (err) console.log(err);
    else console.log(doc);
  });
});

function App = () => {
  return (
    <>
    <h2>Billing</h2>
    <button onclick="showTab('invoice')">Invoices</button>
    <button onclick="showTab('items')">Items</button>
  
    <div id="invoice-tab" class="tab">
      <input id="mobile" placeholder="Mobile">
      <button id="get-cust">Get Customer</button>
      <input id="customer" placeholder="Customer">
      <textarea id="address" placeholder="Address"></textarea>
      <button id="add-cust">Add Customer</button>
      <input id="date" type="date">
      <table id="items">
        <tr><th>Item</th><th>Qty</th><th>Price</th><th>Total</th><th></th></tr>
      </table>
      <button id="add">Add Row</button>
      <button id="save">Save Invoice</button>
      <h3 id="grand-total">Total: ₹0.00</h3>
      <pre id="out"></pre>
      <pre id="cust-out"></pre>
    </div>
  
    <div id="items-tab" class="tab" style="display:none">
      <h3>Manage Items</h3>
      <input id="item-name" placeholder="Item name">
      <input id="item-price" type="number" placeholder="Price">
      <button id="save-item">Add Item</button>
      <ul id="item-list"></ul>
    </div>
    </>
  );
};

ReactDOM.render(<App />, document.getElementById('root'));