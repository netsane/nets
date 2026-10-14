// TestPDFExport.jsx
import React from 'react';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

const TestPDFExport = () => {
  const handleExport = () => {
    const doc = new jsPDF();

    doc.autoTable({
      head: [['Name', 'Email']],
      body: [
        ['John Doe', 'john@example.com'],
        ['Jane Smith', 'jane@example.com'],
      ],
    });

    doc.save('test.pdf');
  };

  return (
    <button onClick={handleExport}>
      Export PDF
    </button>
  );
};

export default TestPDFExport;