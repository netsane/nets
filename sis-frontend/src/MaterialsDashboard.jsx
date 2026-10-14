// MaterialDashboard.jsx
import React from 'react';

const MaterialDashboard = ({ data }) => {
  // Mock data structure: { totalItems: 1250, lowStock: 35, avgLeadTime: 7 }
  const cardData = [
    { title: 'Total SKUs in Inventory', value: data.totalItems.toLocaleString(), className: 'total' },
    { title: 'Low Stock Alerts', value: data.lowStock, className: 'warning' },
    { title: 'Avg. Supplier Lead Time (Days)', value: data.avgLeadTime, className: 'info' },
  ];

  return (
    <div className="dashboard-container">
      <h2>Inventory Overview</h2>
      <div className="card-grid">
        {cardData.map((card, index) => (
          <div key={index} className={`stat-card ${card.className}`}>
            <h3 className="card-title">{card.title}</h3>
            <p className="card-value">{card.value}</p>
          </div>
        ))}
      </div>
      {/* Placeholder for an inventory chart */}
      <div className="chart-placeholder">
        <h3>Inventory Trends Chart</h3>
        <p>A graph showing stock movement over the last 30 days would go here.</p>
      </div>
    </div>
  );
};

export default MaterialDashboard;

/*
// Example CSS Classes (Conceptual)
.dashboard-container { padding: 20px; }
.card-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; }
.stat-card { padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
.stat-card.total { background-color: #e0f7fa; border-left: 5px solid #00bcd4; }
.stat-card.warning { background-color: #fff3e0; border-left: 5px solid #ff9800; }
.stat-card.info { background-color: #ede7f6; border-left: 5px solid #673ab7; }
.card-title { font-size: 14px; color: #555; margin-bottom: 10px; }
.card-value { font-size: 32px; font-weight: bold; }
.chart-placeholder { margin-top: 40px; padding: 20px; border: 1px dashed #ccc; text-align: center; }
*/