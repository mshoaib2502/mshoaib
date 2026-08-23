import React, { useEffect, useState } from 'react';

const Billing = () => {
  const [items, setItems] = useState([]);
  const [customer, setCustomer] = useState('Walk-in');
  const [mobile, setMobile] = useState('');
  const [address, setAddress] = useState('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [grandTotal, setGrandTotal] = useState(0);

  useEffect(() => {
    showTab('invoice');
    loadItemList();
  }, []);

  const showTab = (t) => {
    document.getElementById('invoice-tab').style.display = t === 'invoice' ? 'block' : 'none';
    document.getElementById('items-tab').style.display = t === 'items' ? 'block' : 'none';
  };

  const addRow = () => {
    setItems([...items, { description: '', qty: 1, price: 0 }]);
  };

  const updateRow = (index, field, value) => {
    const newItems = [...items];
    newItems[index][field] = value;
    setItems(newItems);
    updateGrandTotal(newItems);
  };

  const updateGrandTotal = (items) => {
    const total = items.reduce((acc, item) => acc + item.qty * item.price, 0);
    setGrandTotal(total);
  };

  const saveInvoice = async () => {
    if (!window.api) {
      await new Promise(resolve => setTimeout(resolve, 100));
      return;
    }
    const invoiceItems = items.filter(item => item.description);
    await window.api.saveInvoice({ customer, mobile, address, date, items: invoiceItems });
  };

  const loadItemList = async () => {
    if (!window.api) {
      await new Promise(resolve => setTimeout(resolve, 100));
      return loadItemList();
    }
    const list = await window.api.listItems();
    // Handle item list loading logic here
  };

  return (
    <div>
      <div id="invoice-tab">
        {/* Invoice Tab Content */}
        <button onClick={saveInvoice}>Save Invoice</button>
      </div>
      <div id="items-tab">
        <table>
          <tbody>
            {items.map((item, index) => (
              <tr key={index}>
                <td>
                  <select
                    onChange={(e) => updateRow(index, 'description', e.target.value)}
                  >
                    {/* Populate options here */}
                  </select>
                </td>
                <td>
                  <input
                    type="number"
                    value={item.qty}
                    onChange={(e) => updateRow(index, 'qty', parseFloat(e.target.value))}
                  />
                </td>
                <td>
                  <input
                    type="number"
                    value={item.price}
                    onChange={(e) => updateRow(index, 'price', parseFloat(e.target.value))}
                  />
                </td>
                <td>₹{(item.qty * item.price).toFixed(2)}</td>
                <td>
                  <button onClick={() => setItems(items.filter((_, i) => i !== index))}>x</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <button onClick={addRow}>Add Item</button>
        <div>Total: ₹{grandTotal.toFixed(2)}</div>
      </div>
    </div>
  );
};

export default Billing;
