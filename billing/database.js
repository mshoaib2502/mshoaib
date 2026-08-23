const Database = require('better-sqlite3');
const db = new Database('billing.db');

function initDB() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS products (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT, hsn TEXT, price REAL, gst REAL, stock INTEGER
    );
    CREATE TABLE IF NOT EXISTS customers (
      phone TEXT PRIMARY KEY,
      name TEXT NOT NULL, 
      gstin TEXT, 
      address TEXT
    );
    CREATE TABLE IF NOT EXISTS invoices (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer TEXT, customer_phone TEXT, items TEXT, total REAL, gst_total REAL, date TEXT
    );
  `);
}

// Products
function getProducts() { return db.prepare('SELECT * FROM products').all(); }
function addProduct(p) { return db.prepare('INSERT INTO products (name,hsn,price,gst,stock) VALUES (?,?,?,?,?)').run(p.name, p.hsn, p.price, p.gst, p.stock); }
function deleteProduct(id) { return db.prepare('DELETE FROM products WHERE id=?').run(id); }

// Customers - PHONE IS ID NOW
function getCustomers() { return db.prepare('SELECT * FROM customers ORDER BY name').all(); }
function addCustomer(c) { 
  if(!c.phone || !c.name) throw new Error('Phone and Name required');
  return db.prepare('INSERT OR REPLACE INTO customers (phone, name, gstin, address) VALUES (?,?,?,?)').run(c.phone, c.name, c.gstin || '', c.address || ''); 
}
function searchCustomers(q) { return db.prepare('SELECT * FROM customers WHERE name LIKE ? OR phone LIKE ?').all(`%${q}%`, `%${q}%`); }

// Invoices
function createInvoice(inv) { return db.prepare('INSERT INTO invoices (customer, customer_phone, items, total, gst_total, date) VALUES (?,?,?,?,?,?)').run(inv.customer, inv.customer_phone, JSON.stringify(inv.items), inv.total, inv.gst_total, new Date().toISOString()); }
function getInvoices() { return db.prepare('SELECT * FROM invoices ORDER BY id DESC').all(); }

module.exports = { initDB, getProducts, addProduct, deleteProduct, getCustomers, addCustomer, searchCustomers, createInvoice, getInvoices };