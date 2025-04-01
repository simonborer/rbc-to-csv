// Create a CSV from your online banking
// Load all the transactions, then run this in the console

// Yes, RBC says they'll do this for you, but they only
// let you download the last couple of months!
(() => {
  const normalizeAmount = (amount) =>
    amount
      .replace(/\$/, '')
      .replace(/-/, '')
      .replace(/\s/g, '')
      .replace(',', '.');

  const extractDateFromId = (id) => {
    const match = id.match(/\d{4}-\d{2}-\d{2}/); // extract the date pattern
    if (!match) {
      throw new Error(`Could not extract a valid date from id "${id}".`);
    }
    return match[0]; // return just the date part
  };

  const createDownloadLink = (linkElement, content) => {
    if (linkElement?.nodeName !== 'A') {
      throw new Error(`Expected an anchor <a> element for download.`);
    }
    const blob = new Blob([content], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    linkElement.href = url;
    linkElement.download = 'transactions.csv';
    linkElement.innerHTML = '<span>CSV</span>';
  };

  const extractTransactions = (table) => {
    if (table?.nodeName !== 'TABLE') {
      throw new Error(`Expected a <table> element with transaction data.`);
    }

    const rows = Array.from(table.querySelectorAll('tbody tr'));
    const csvLines = [`"Date","Description","Transaction","Debit","Credit","Total"`];

    rows.forEach((row) => {
      const cells = row.querySelectorAll('td');
      if (cells.length < 5) return; // Skip malformed rows

      const [dateCell, descCell, debitCell, creditCell, totalCell] = cells;

      const date = extractDateFromId(dateCell.id || '');
      const description = (descCell.textContent || '').trim();
      const transactionId = /\s-\s(\d+)/.exec(description)?.[1] || '';
      const debit = normalizeAmount(debitCell.textContent || '');
      const credit = normalizeAmount(creditCell.textContent || '');
      const total = normalizeAmount(totalCell.textContent || '');

      const safeDesc = description.replace(/"/g, '""'); // escape quotes for CSV

      csvLines.push(`"${date}","${safeDesc}","${transactionId}","${debit}","${credit}","${total}"`);
    });

    return csvLines.join('\n');
  };

  // === Execute ===
  const table = document.querySelector('table.rbc-transaction-list-table');
  const downloadButton = document.querySelector('[rbcportalsubmit="DownloadTransactions"]');

  const csvContent = extractTransactions(table);
  createDownloadLink(downloadButton, csvContent);
})();
