(() => {
  const normalizeAmount = (amount) =>
    amount
      .replace(/\$/, '')
      .replace(/-/, '')
      .replace(/\s/g, '')
      .replace(',', '.');

  const extractDateFromId = (id) => {
    const match = id.match(/\d{4}-\d{2}-\d{2}/);
    if (!match) {
      throw new Error(`Could not extract a valid date from id "${id}".`);
    }
    return match[0];
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

  const findTableWithMostRows = () => {
    const tables = document.querySelectorAll('table.rbc-transaction-list-table');
    if (!tables.length) {
      throw new Error('No transaction tables found.');
    }

    let bestTable = null;
    let maxRows = 0;

    tables.forEach((table) => {
      const rowCount = table.querySelectorAll('tbody tr').length;
      console.log('Table found with', rowCount, 'rows.');
      if (rowCount > maxRows) {
        maxRows = rowCount;
        bestTable = table;
      }
    });

    if (!bestTable) {
      throw new Error('Could not find a suitable transaction table.');
    }

    console.log('Selected table with', maxRows, 'rows.');
    return bestTable;
  };

  const extractTransactions = (table) => {
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    console.log('Extracting', rows.length, 'rows from table.');

    const csvLines = [`"Date","Description","Transaction","Debit","Credit","Total"`];

    rows.forEach((row, index) => {
      const cells = row.querySelectorAll('th, td'); // <-- FIXED
      if (cells.length < 4) {
        console.warn(`Skipping row ${index + 1} due to insufficient cells.`);
        return;
      }

      const [dateCell, descCell, debitCell, creditCell, totalCell] = cells;

      const date = extractDateFromId(dateCell.id || '');
      const description = (descCell.textContent || '').trim();
      const transactionId = /\s-\s(\d+)/.exec(description)?.[1] || '';
      const debit = normalizeAmount(debitCell.textContent || '');
      const credit = normalizeAmount(creditCell.textContent || '');
      const total = cells.length >= 5 ? normalizeAmount(totalCell.textContent || '') : '';

      const safeDesc = description.replace(/"/g, '""');

      csvLines.push(`"${date}","${safeDesc}","${transactionId}","${debit}","${credit}","${total}"`);
    });

    return csvLines.join('\n');
  };

  // === Execute ===
  const table = findTableWithMostRows();
  const downloadButton = document.querySelector('[rbcportalsubmit="DownloadTransactions"]');

  if (!downloadButton) {
    throw new Error('Download button not found.');
  }

  const csvContent = extractTransactions(table);
  createDownloadLink(downloadButton, csvContent);

  console.log('CSV download link ready.');
})();
