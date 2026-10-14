// EmployeePayslip.jsx
import React from 'react';

const EmployeePayslip = ({ payslip }) => {
  // Mock payslip structure: 
  // { employeeName: 'Jane Doe', period: 'Dec 1 - Dec 31, 2024', baseSalary: 6000, 
  //   earnings: [{type: 'Bonus', amount: 500}], 
  //   deductions: [{type: 'Federal Tax', amount: 1200}, {type: 'Health Ins.', amount: 150}] 
  // }
  
  const totalEarnings = payslip.baseSalary + payslip.earnings.reduce((sum, e) => sum + e.amount, 0);
  const totalDeductions = payslip.deductions.reduce((sum, d) => sum + d.amount, 0);
  const netPay = totalEarnings - totalDeductions;

  const formatCurrency = (amount) => amount.toLocaleString('en-US', { style: 'currency', currency: 'USD' });

  return (
    <div className="payslip-container">
      <header className="payslip-header">
        <h1>Official Payslip</h1>
        <p>Pay Period: **{payslip.period}**</p>
      </header>
      
      <section className="employee-info">
        <h3>Employee Details</h3>
        <p>Name: **{payslip.employeeName}**</p>
        <p>Employee ID: **{payslip.employeeId || 'EMP-1234'}**</p>
        <p>Department: **{payslip.department || 'Finance'}**</p>
      </section>

      <div className="payslip-details-grid">
        {/* Earnings Section */}
        <div className="section earnings-section">
          <h3>Earnings</h3>
          <div className="detail-row">
            <span>Base Salary</span>
            <span className="amount">{formatCurrency(payslip.baseSalary)}</span>
          </div>
          {payslip.earnings.map((e, index) => (
            <div key={index} className="detail-row">
              <span>{e.type}</span>
              <span className="amount">{formatCurrency(e.amount)}</span>
            </div>
          ))}
          <div className="detail-row total-line">
            <span>**TOTAL EARNINGS**</span>
            <span className="amount">**{formatCurrency(totalEarnings)}**</span>
          </div>
        </div>

        {/* Deductions Section */}
        <div className="section deductions-section">
          <h3>Deductions</h3>
          {payslip.deductions.map((d, index) => (
            <div key={index} className="detail-row">
              <span>{d.type}</span>
              <span className="amount">({formatCurrency(d.amount)})</span>
            </div>
          ))}
          <div className="detail-row total-line">
            <span>**TOTAL DEDUCTIONS**</span>
            <span className="amount">**({formatCurrency(totalDeductions)})**</span>
          </div>
        </div>
      </div>

      <footer className="payslip-footer">
        <div className="net-pay-box">
          <span>NET PAY</span>
          <span className="net-pay-amount">**{formatCurrency(netPay)}**</span>
        </div>
        <button onClick={() => window.print()} className="print-button">
          Print Payslip
        </button>
      </footer>
    </div>
  );
};

export default EmployeePayslip;

/*
// Example CSS Classes (Conceptual)
.payslip-container { max-width: 600px; margin: 30px auto; padding: 30px; border: 1px solid #ccc; font-family: sans-serif; }
.payslip-header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 20px; }
.employee-info { margin-bottom: 20px; padding: 10px; border: 1px dashed #eee; }
.payslip-details-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.section h3 { background-color: #f5f5f5; padding: 8px; margin-bottom: 10px; }
.detail-row { display: flex; justify-content: space-between; padding: 4px 0; border-bottom: 1px dotted #eee; }
.total-line { font-weight: bold; border-top: 2px solid #333; margin-top: 8px; padding-top: 8px; }
.payslip-footer { margin-top: 30px; text-align: center; }
.net-pay-box { background-color: #d1e7dd; color: #0a3622; padding: 15px; border-radius: 5px; display: flex; justify-content: space-between; font-size: 1.5em; font-weight: bold; margin-bottom: 15px; }
.print-button { padding: 10px 20px; cursor: pointer; }
*/