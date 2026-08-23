import React from 'react';
import { contextBridge, ipcRenderer } from 'electron';

const App = () => {
  const saveInvoice = (data) => ipcRenderer.invoke('save-invoice', data);
  const listInvoices = () => ipcRenderer.invoke('list-invoices');
  const addItem = (data) => ipcRenderer.invoke('add-item', data);
  const listItems = () => ipcRenderer.invoke('list-items');
  const getCustomer = (mobile) => ipcRenderer.invoke('get-customer', mobile);
  const saveCustomer = (data) => ipcRenderer.invoke('save-customer', data);

  return (
    <div>
      <h1>Invoice Management</h1>
      {/* Add your UI components and event handlers here */}
    </div>
  );
};

export default App;