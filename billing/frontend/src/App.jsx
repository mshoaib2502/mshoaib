import { useState, useEffect } from 'react';

export default function App() {
  const [tab, setTab] = useState('billing');
  const [products, setProducts] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [cart, setCart] = useState([]);
  
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [billForm, setBillForm] = useState({ name: '', price: '', qty: '1', gst: '18' });
  const [prodForm, setProdForm] = useState({ name: '', hsn: '', price: '', gst: '18', stock: '' });
  const [custForm, setCustForm] = useState({ name: '', phone: '', gstin: '', address: '' });

  useEffect(() => { loadAll(); }, []);
  async function loadAll() {
    if (!window.api) return;
    setProducts(await window.api.getProducts());
    setCustomers(await window.api.getCustomers());
    setInvoices(await window.api.getInvoices());
  }

  const total = cart.reduce((s, i) => s + i.price * i.qty, 0);
  const gstTotal = cart.reduce((s, i) => s + (i.price * i.qty * i.gst / 100), 0);

  return (
    <div style={{ fontFamily: 'system-ui' }}>
      <div style={{ display: 'flex', gap: 10, padding: 12, background: '#111' }}>
        {['billing','products','customers','invoices'].map(t => (
          <button key={t} onClick={()=>{ setTab(t); loadAll(); }} style={{ padding:'8px 16px', background:tab===t?'white':'#333', color:tab===t?'black':'white', border:0, borderRadius:6, textTransform:'capitalize' }}>{t}</button>
        ))}
      </div>

      {tab==='billing' && (
        <div style={{ padding:20 }}>
          <h3>Billing - Add Products for Customer Bill</h3>
          <div style={{ display:'flex', gap:10, marginBottom:15 }}>
            <input placeholder="Customer Phone *" value={customerPhone} onChange={e=>setCustomerPhone(e.target.value)} style={{ padding:10, width:180 }} />
            <input placeholder="Customer Name *" value={customerName} onChange={e=>setCustomerName(e.target.value)} style={{ padding:10, flex:1 }} />
          </div>

          <div style={{ background:'#f9f9f9', padding:12, borderRadius:8, marginBottom:15 }}>
            <form onSubmit={e=>{ e.preventDefault(); if(!billForm.name || !billForm.price) return; setCart([...cart, { name:billForm.name, price:+billForm.price, qty:+(billForm.qty||1), gst:+(billForm.gst||0), hsn:'9988' }]); setBillForm({ name:'', price:'', qty:'1', gst:'18' }); }} style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
              <input required list="prodMaster" placeholder="Product Name" value={billForm.name} onChange={e=>setBillForm({...billForm, name:e.target.value})} style={{ padding:8, flex:1, minWidth:150 }} />
              <datalist id="prodMaster">{products.map(p=><option key={p.id} value={p.name}>₹{p.price}</option>)}</datalist>
              <input required placeholder="Price ₹" type="number" value={billForm.price} onChange={e=>setBillForm({...billForm, price:e.target.value})} style={{ padding:8, width:100 }} />
              <input placeholder="Qty" type="number" value={billForm.qty} onChange={e=>setBillForm({...billForm, qty:e.target.value})} style={{ padding:8, width:70 }} />
              <select value={billForm.gst} onChange={e=>setBillForm({...billForm, gst:e.target.value})} style={{ padding:8 }}><option>0</option><option>5</option><option>12</option><option>18</option><option>28</option></select>
              <button type="submit" style={{ background:'black', color:'white', border:0, padding:'8px 16px', borderRadius:6 }}>Add to Bill</button>
            </form>
            {products.length>0 && <div style={{ marginTop:10, display:'flex', gap:6, flexWrap:'wrap' }}>{products.slice(0,8).map(p=><button key={p.id} onClick={()=>setBillForm({ name:p.name, price:p.price.toString(), qty:'1', gst:p.gst.toString() })} style={{ padding:'4px 10px', border:'1px solid #ddd', background:'white', borderRadius:12, fontSize:12 }}>{p.name} ₹{p.price}</button>)}</div>}
          </div>

          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead><tr style={{ background:'#f5f5f5' }}><th style={{ border:'1px solid #ddd', padding:8 }}>Product</th><th style={{ border:'1px solid #ddd', padding:8 }}>Price</th><th style={{ border:'1px solid #ddd', padding:8 }}>Qty</th><th style={{ border:'1px solid #ddd', padding:8 }}>Amount</th><th style={{ border:'1px solid #ddd', padding:8 }}></th></tr></thead>
            <tbody>{cart.map((c,i)=><tr key={i}><td style={{ border:'1px solid #ddd', padding:8 }}>{c.name}</td><td style={{ border:'1px solid #ddd', padding:8 }}>₹{c.price}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{c.qty}</td><td style={{ border:'1px solid #ddd', padding:8 }}>₹{(c.price*c.qty*(1+c.gst/100)).toFixed(2)}</td><td style={{ border:'1px solid #ddd', padding:8 }}><button onClick={()=>setCart(cart.filter((_,idx)=>idx!==i))}>Remove</button></td></tr>)}</tbody>
          </table>
          <div style={{ textAlign:'right', marginTop:15 }}><h2>Total: ₹{(total+gstTotal).toFixed(2)}</h2>
            <button onClick={async()=>{
              if(!customerPhone || !customerName) return alert('Enter Phone & Name');
              if(cart.length===0) return alert('Add product');
              await window.api.addCustomer({ phone: customerPhone, name: customerName, gstin:'', address:'' });
              await window.api.createInvoice({ customer: customerName, customer_phone: customerPhone, items: cart, total: total+gstTotal, gst_total: gstTotal });
              setCart([]); setCustomerName(''); setCustomerPhone(''); loadAll(); alert('Bill saved & Customer added!');
            }} style={{ padding:12, background:'black', color:'white', border:0, width:'100%' }}>Save Bill</button>
          </div>
        </div>
      )}

      {tab==='products' && (
        <div style={{ padding:20 }}>
          <h3>Shop Master - Maintain All Products</h3>
          <form onSubmit={async e=>{
            e.preventDefault();
            await window.api.addProduct({ name:prodForm.name, hsn:prodForm.hsn||'9988', price:+prodForm.price, gst:+prodForm.gst, stock:+(prodForm.stock||0) });
            setProdForm({ name:'', hsn:'', price:'', gst:'18', stock:'' }); loadAll();
          }} style={{ display:'flex', gap:8, flexWrap:'wrap', marginBottom:20, background:'#f9f9f9', padding:12, borderRadius:8 }}>
            <input required placeholder="Product Name *" value={prodForm.name} onChange={e=>setProdForm({...prodForm, name:e.target.value})} style={{ padding:8, minWidth:150 }} />
            <input placeholder="HSN" value={prodForm.hsn} onChange={e=>setProdForm({...prodForm, hsn:e.target.value})} style={{ padding:8, width:90 }} />
            <input required placeholder="Price *" type="number" value={prodForm.price} onChange={e=>setProdForm({...prodForm, price:e.target.value})} style={{ padding:8, width:110 }} />
            <select value={prodForm.gst} onChange={e=>setProdForm({...prodForm, gst:e.target.value})} style={{ padding:8 }}><option value="0">0%</option><option value="5">5%</option><option value="12">12%</option><option value="18">18%</option><option value="28">28%</option></select>
            <input placeholder="Stock" type="number" value={prodForm.stock} onChange={e=>setProdForm({...prodForm, stock:e.target.value})} style={{ padding:8, width:90 }} />
            <button style={{ background:'black', color:'white', border:0, padding:'8px 16px' }}>Add to Master</button>
          </form>

          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead><tr style={{ background:'#f5f5f5' }}><th style={{ border:'1px solid #ddd', padding:8 }}>Name</th><th style={{ border:'1px solid #ddd', padding:8 }}>HSN</th><th style={{ border:'1px solid #ddd', padding:8 }}>Price</th><th style={{ border:'1px solid #ddd', padding:8 }}>GST</th><th style={{ border:'1px solid #ddd', padding:8 }}>Stock</th><th style={{ border:'1px solid #ddd', padding:8 }}>Action</th></tr></thead>
            <tbody>{products.map(p=><tr key={p.id}><td style={{ border:'1px solid #ddd', padding:8 }}>{p.name}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{p.hsn}</td><td style={{ border:'1px solid #ddd', padding:8 }}>₹{p.price}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{p.gst}%</td><td style={{ border:'1px solid #ddd', padding:8 }}>{p.stock}</td><td style={{ border:'1px solid #ddd', padding:8 }}><button onClick={async()=>{ if(confirm('Delete?')){ await window.api.deleteProduct(p.id); loadAll(); }}}>Delete</button></td></tr>)}</tbody>
          </table>
          {products.length===0 && <p style={{ color:'#666' }}>No products in master yet. Add one above.</p>}
        </div>
      )}

      {tab==='customers' && (
        <div style={{ padding:20 }}>
          <h3>Customers - Phone</h3>
          <form onSubmit={async e=>{
            e.preventDefault();
            if(!custForm.phone || !custForm.name) return alert('Phone and Name required');
            await window.api.addCustomer(custForm);
            setCustForm({ name:'', phone:'', gstin:'', address:'' });
            loadAll(); alert('Customer saved');
          }} style={{ display:'flex', gap:8, flexWrap:'wrap', marginBottom:20, background:'#f9f9f9', padding:12 }}>
            <input required placeholder="Phone *" value={custForm.phone} onChange={e=>setCustForm({...custForm, phone:e.target.value})} style={{ padding:8 }} />
            <input required placeholder="Name *" value={custForm.name} onChange={e=>setCustForm({...custForm, name:e.target.value})} style={{ padding:8 }} />
            <input placeholder="GSTIN" value={custForm.gstin} onChange={e=>setCustForm({...custForm, gstin:e.target.value})} style={{ padding:8 }} />
            <input placeholder="Address" value={custForm.address} onChange={e=>setCustForm({...custForm, address:e.target.value})} style={{ padding:8, width:200 }} />
            <button style={{ background:'black', color:'white', border:0, padding:'8px 16px' }}>Add / Update Customer</button>
          </form>
          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead><tr style={{ background:'#f5f5f5' }}><th style={{ border:'1px solid #ddd', padding:8 }}>Phone</th><th style={{ border:'1px solid #ddd', padding:8 }}>Name</th><th style={{ border:'1px solid #ddd', padding:8 }}>GSTIN</th><th style={{ border:'1px solid #ddd', padding:8 }}>Address</th></tr></thead>
            <tbody>{customers.map(c=><tr key={c.phone}><td style={{ border:'1px solid #ddd', padding:8 }}><b>{c.phone}</b></td><td style={{ border:'1px solid #ddd', padding:8 }}>{c.name}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{c.gstin}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{c.address}</td></tr>)}</tbody>
          </table>
        </div>
      )}

      {tab==='invoices' && (
        <div style={{ padding:20 }}>
          <h3>Invoices History</h3>
          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead><tr style={{ background:'#f5f5f5' }}><th style={{ border:'1px solid #ddd', padding:8 }}>ID</th><th style={{ border:'1px solid #ddd', padding:8 }}>Customer</th><th style={{ border:'1px solid #ddd', padding:8 }}>Phone</th><th style={{ border:'1px solid #ddd', padding:8 }}>Products Billed</th><th style={{ border:'1px solid #ddd', padding:8 }}>Total</th><th style={{ border:'1px solid #ddd', padding:8 }}>Date</th></tr></thead>
            <tbody>{invoices.map(inv=>{ const items=JSON.parse(inv.items||'[]'); return <tr key={inv.id}><td style={{ border:'1px solid #ddd', padding:8 }}>{inv.id}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{inv.customer}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{inv.customer_phone}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{items.map(i=>`${i.name} x${i.qty} @₹${i.price}`).join(', ')}</td><td style={{ border:'1px solid #ddd', padding:8 }}>₹{inv.total?.toFixed(2)}</td><td style={{ border:'1px solid #ddd', padding:8 }}>{new Date(inv.date).toLocaleString()}</td></tr>})}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}