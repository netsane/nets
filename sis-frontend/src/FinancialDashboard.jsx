// TrialBalanceTable.jsx
import React from 'react';

const TrialBalanceTable = ({ accounts }) => {
  // Mock data structure: [{id: 1, name: 'Cash', debit: 50000, credit: 0}, ...]
  
  const totalDebit = accounts.reduce((sum, acc) => sum + acc.debit, 0);
  const totalCredit = accounts.reduce((sum, acc) => sum + acc.credit, 0);

  return (
    <div className="report-container">
      <h2>Trial Balance Report</h2>
      <p className="report-period">As of: December 31, 2024</p>
      
      <table className="trial-balance-table">
        <thead>
          <tr>
            <th>Account Name</th>
            <th className="align-right">Debit ($)</th>
            <th className="align-right">Credit ($)</th>
          </tr>
        </thead>
        <tbody>
          {accounts.map((account) => (
            <tr key={account.id} className="account-row">
              <td>{account.name}</td>
              <td className="align-right">{account.debit.toLocaleString('en-US')}</td>
              <td className="align-right">{account.credit.toLocaleString('en-US')}</td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr className={`total-row ${totalDebit === totalCredit ? 'balanced' : 'unbalanced'}`}>
            <td className="total-label">TOTALS</td>
            <td className="align-right total-amount">{totalDebit.toLocaleString('en-US')}</td>
            <td className="align-right total-amount">{totalCredit.toLocaleString('en-US')}</td>
          </tr>
        </tfoot>
      </table>

      {totalDebit !== totalCredit && (
        <p className="balance-alert">
          🚨 **Warning:** The Trial Balance is **UNBALANCED**. Debit and Credit totals do not match.
        </p>
      )}
    </div>
  );
};

export default TrialBalanceTable;

/*
// Example CSS Classes (Conceptual)
.trial-balance-table { width: 100%; border-collapse: collapse; margin-top: 15px; }
.trial-balance-table th, .trial-balance-table td { border: 1px solid #ddd; padding: 12px; }
.trial-balance-table thead th { background-color: #f2f2f2; text-align: left; }
.align-right { text-align: right; }
.total-row { font-weight: bold; }
.total-row.balanced { background-color: #e6ffe6; }
.total-row.unbalanced { background-color: #ffe6e6; }
.balance-alert { color: red; font-weight: bold; margin-top: 15px; padding: 10px; border: 1px solid red; background-color: #fdd; }
*/