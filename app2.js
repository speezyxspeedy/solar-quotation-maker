import { useState } from "react";

export default function InvoiceGenerator() {
  const [items, setItems] = useState([
    { desc: "SOLAR PANEL", hsn: "85414300", qty: 1, unit: "NOS", rate: 0 }
  ]);

  const company = {
    name: "SHREE GANESH SOLAR ENERGY",
    address:
      "Ground, Plot No 1764/29, C/O Mitaram Bhude, Bhandewadi Pardi, Near Hanuman Mandir, Hanuman Nagar, Nagpur",
    gstin: "27DSQPB6025M1ZC",
    state: "Maharashtra, Code: 27"
  };

  const subtotal = items.reduce((s, i) => s + Number(i.qty) * Number(i.rate), 0);
  const cgst = subtotal * 0.09;
  const sgst = subtotal * 0.09;
  const total = Math.round(subtotal + cgst + sgst);

  const updateItem = (index, key, value) => {
    const copy = [...items];
    copy[index][key] = value;
    setItems(copy);
  };

  const addRow = () => {
    setItems([...items, { desc: "", hsn: "", qty: 1, unit: "NOS", rate: 0 }]);
  };

  const deleteRow = (index) => {
    setItems(items.filter((_, i) => i !== index));
  };

  return (
    <div style={{ padding: 20 }}>
      <h1>SHREE GANESH SOLAR ENERGY</h1>
      <h2>Solar Tax Invoice Generator</h2>

      <div style={{ border: "1px solid #000", padding: 15 }}>
        <h3>Seller</h3>
        <b>{company.name}</b>
        <p>{company.address}</p>
        <p>GSTIN/UIN: {company.gstin}</p>
        <p>State: {company.state}</p>
      </div>

      <br />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <input placeholder="Invoice No" />
        <input type="date" />
        <input placeholder="Customer Name" />
        <input placeholder="Customer GSTIN" />
        <textarea placeholder="Bill To Address" />
        <textarea placeholder="Ship To Address" />
        <input placeholder="Payment Terms" defaultValue="Immediate" />
        <input placeholder="Destination" defaultValue="Nagpur" />
      </div>

      <br />

      <table border="1" width="100%" cellPadding="8">
        <thead>
          <tr>
            <th>Sl</th>
            <th>Description</th>
            <th>HSN/SAC</th>
            <th>Qty</th>
            <th>Unit</th>
            <th>Rate</th>
            <th>Amount</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              <td>
                <input
                  value={item.desc}
                  onChange={(e) => updateItem(i, "desc", e.target.value)}
                />
              </td>
              <td>
                <input
                  value={item.hsn}
                  onChange={(e) => updateItem(i, "hsn", e.target.value)}
                />
              </td>
              <td>
                <input
                  type="number"
                  value={item.qty}
                  onChange={(e) => updateItem(i, "qty", e.target.value)}
                />
              </td>
              <td>
                <input
                  value={item.unit}
                  onChange={(e) => updateItem(i, "unit", e.target.value)}
                />
              </td>
              <td>
                <input
                  type="number"
                  value={item.rate}
                  onChange={(e) => updateItem(i, "rate", e.target.value)}
                />
              </td>
              <td>₹{Number(item.qty) * Number(item.rate)}</td>
              <td>
                <button onClick={() => deleteRow(i)}>Delete</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <br />
      <button onClick={addRow}>+ Add Row</button>

      <h3>Subtotal: ₹{subtotal.toFixed(2)}</h3>
      <h3>CGST 9%: ₹{cgst.toFixed(2)}</h3>
      <h3>SGST 9%: ₹{sgst.toFixed(2)}</h3>
      <h2>Total: ₹{total}</h2>

      <button onClick={() => window.print()}>Print / Download PDF</button>
    </div>
  );
}