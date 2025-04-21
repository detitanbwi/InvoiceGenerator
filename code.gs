function createInvoiceFromTemplate(templateId, invoiceData, newFileName) {
  try {
    // 1. Duplicate the template file
    var templateFile = DriveApp.getFileById(templateId);
    var copiedFile = templateFile.makeCopy();
    var copiedFileId = copiedFile.getId();
    var copiedDoc = DocumentApp.openById(copiedFileId);
    var body = copiedDoc.getBody();

    // 2. Modify and fill the format in the duplicated document
    // *** You need to adjust the code here according to your template format ***
    // Example: Replace placeholder text with invoice data
    body.replaceText('{{invoice_number}}', 'INV-2025-001'); // Example static invoice number
    body.replaceText('{{invoice_date}}', Utilities.formatDate(new Date(), Session.getTimeZone(), 'dd-MM-yyyy'));

    // Assuming you have a function to populate item details in the table
    populateInvoiceTable(body, invoiceData);

    copiedDoc.saveAndClose();

    // 3. Rename the file
    copiedFile.setName(newFileName);

    // 4. Convert to PDF
    var pdfBlob = copiedFile.getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);

    // 5. Delete the original Google Docs file (duplicate)
    DriveApp.getFileById(copiedFileId).setTrashed(true);
    Logger.log('Google Docs file with ID ' + copiedFileId + ' has been deleted (moved to trash).');
    Logger.log('PDF file with name ' + newFileName + '.pdf has been created with ID ' + pdfFile.getId());

  } catch (error) {
    Logger.log('An error occurred: ' + error);
  }
}

function populateInvoiceTable(body, invoiceData) {
  // *** This function needs to be adjusted according to your invoice table structure ***
  // Simple example if the invoice table is the first table in the document:
  var tables = body.getTables();
  if (tables.length > 0) {
    var invoiceTable = tables[0]; // Assuming the invoice table is the first table

    Logger.log("Number of rows before deletion: " + invoiceTable.getNumRows());

    // Delete existing item data rows (if any, excluding the header)
    // Ensure there is more than one row before attempting to delete
    if (invoiceTable.getNumRows() > 1) {
      // Delete rows from bottom to top to avoid index changes
      for (var i = invoiceTable.getNumRows() - 1; i > 0; i--) {
        invoiceTable.removeRow(i); // Using removeRow with index
      }
    } else {
      Logger.log("Only one row (header) or no rows exist, no need to delete.");
    }

    // Add new rows based on invoice data
    for (var i = 0; i < invoiceData.length; i++) {
      var item = invoiceData[i];
      var newRow = invoiceTable.appendTableRow();
      newRow.appendTableCell(item.name).setText(item.name);
      newRow.appendTableCell(item.quantity.toString()).setText(item.quantity.toString());
      newRow.appendTableCell(item.price.toLocaleString('id-ID')).setText(item.price.toLocaleString('id-ID'));
      newRow.appendTableCell((item.quantity * item.price).toLocaleString('id-ID')).setText((item.quantity * item.price).toLocaleString('id-ID'));
    }

    // Add total row (simple example, needs adjustment)
    var totalRow = invoiceTable.appendTableRow();
    totalRow.appendTableCell('Total');
    totalRow.appendTableCell(''); // Empty cell for quantity column
    totalRow.appendTableCell(''); // Empty cell for unit price column
    totalRow.appendTableCell(invoiceData.reduce(function(sum, item) {
      return sum + (item.quantity * item.price);
    }, 0).toLocaleString('id-ID')).setText(invoiceData.reduce(function(sum, item) {
      return sum + (item.quantity * item.price);
    }, 0).toLocaleString('id-ID'));

    // Merge the first three cells in the total row
    totalRow.getCell(1).merge(); // Merge the second cell with the first
    totalRow.getCell(2).merge(); // Merge the third cell with the second (which is already merged)

  } else {
    Logger.log('Invoice table not found in the document.');
  }
}

// Example usage:
function main() {
  var templateId = 'YOUR TEMPLATE DOCS FILEID HERE'; // Replace with your template file ID
  var invoiceData = [
    { name: "Jasa Desain", quantity: 1, price: 500000 },
    { name: "Cetak Brosur", quantity: 100, price: 5000 },
    { name: "Konsultasi", quantity: 2, price: 250000 }
  ];
  var newFileName = 'Invoice PT Maju Jaya - ' + Utilities.formatDate(new Date(), Session.getTimeZone(), 'yyyyMMdd');

  createInvoiceFromTemplate(templateId, invoiceData, newFileName);
}
