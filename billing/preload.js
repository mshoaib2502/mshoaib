const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('api', {
    getProducts: () => ipcRenderer.invoke('get-products'),
    addProduct: (p) => ipcRenderer.invoke('add-product', p),
    deleteProduct: (id) => ipcRenderer.invoke('delete-product', id),
    getCustomers: () => ipcRenderer.invoke('get-customers'),
    addCustomer: (c) => ipcRenderer.invoke('add-customer', c),
    searchCustomers: (q) => ipcRenderer.invoke('search-customers', q),
    createInvoice: (inv) => ipcRenderer.invoke('create-invoice', inv),
    getInvoices: () => ipcRenderer.invoke('get-invoices'),
  });