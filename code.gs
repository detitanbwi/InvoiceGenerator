function buatInvoiceDariTemplate(templateId, dataInvoice, namaFileBaru) {
  try {
    // 1. Duplikasi file template
    var templateFile = DriveApp.getFileById(templateId);
    var copiedFile = templateFile.makeCopy();
    var copiedFileId = copiedFile.getId();
    var copiedDoc = DocumentApp.openById(copiedFileId);
    var body = copiedDoc.getBody();

    // 2. Mengubah dan mengisi format di dokumen duplikat
    // *** Di sini Anda perlu menyesuaikan kode sesuai dengan format template Anda ***
    // Contoh: Mengganti placeholder teks dengan data invoice
    body.replaceText('{{nomor_invoice}}', 'INV-2025-001'); // Contoh nomor invoice statis
    body.replaceText('{{tanggal_invoice}}', Utilities.formatDate(new Date(), Session.getTimeZone(), 'dd-MM-yyyy'));

    // Asumsi Anda memiliki fungsi untuk mengisi detail item dalam tabel
    isiTabelInvoice(body, dataInvoice);

    copiedDoc.saveAndClose();

    // 3. Mengganti nama file
    copiedFile.setName(namaFileBaru);

    // 4. Mengonversi ke PDF
    var pdfBlob = copiedFile.getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);

    // 5. Menghapus file Google Docs asli (duplikat)
    DriveApp.getFileById(copiedFileId).setTrashed(true);
    Logger.log('File Google Docs dengan ID ' + copiedFileId + ' telah dihapus (dipindahkan ke sampah).');
    Logger.log('File PDF dengan nama ' + namaFileBaru + '.pdf telah dibuat dengan ID ' + pdfFile.getId());

  } catch (error) {
    Logger.log('Terjadi kesalahan: ' + error);
  }
}

function isiTabelInvoice(body, dataInvoice) {
  // *** Fungsi ini perlu disesuaikan dengan struktur tabel invoice Anda ***
  // Contoh sederhana jika tabel invoice adalah tabel pertama dalam dokumen:
  var tables = body.getTables();
  if (tables.length > 0) {
    var invoiceTable = tables[0]; // Asumsi tabel invoice adalah tabel pertama

    Logger.log("Jumlah baris sebelum penghapusan: " + invoiceTable.getNumRows());

    // Hapus baris data item yang ada (jika ada selain header)
    // Pastikan ada lebih dari satu baris sebelum mencoba menghapus
    if (invoiceTable.getNumRows() > 1) {
      // Hapus baris dari bawah ke atas untuk menghindari perubahan indeks
      for (var i = invoiceTable.getNumRows() - 1; i > 0; i--) {
        invoiceTable.removeRow(i); // Menggunakan removeRow dengan indeks
      }
    } else {
      Logger.log("Hanya ada satu baris (header) atau tidak ada baris, tidak perlu menghapus.");
    }

    // Tambahkan baris baru berdasarkan data invoice
    for (var i = 0; i < dataInvoice.length; i++) {
      var item = dataInvoice[i];
      var newRow = invoiceTable.appendTableRow();
      newRow.appendTableCell(item.nama).setText(item.nama);
      newRow.appendTableCell(item.kuantitas.toString()).setText(item.kuantitas.toString());
      newRow.appendTableCell(item.harga.toLocaleString('id-ID')).setText(item.harga.toLocaleString('id-ID'));
      newRow.appendTableCell((item.kuantitas * item.harga).toLocaleString('id-ID')).setText((item.kuantitas * item.harga).toLocaleString('id-ID'));
    }

    // Tambahkan baris total (contoh sederhana, perlu disesuaikan)
    var totalRow = invoiceTable.appendTableRow();
    totalRow.appendTableCell('Total');
    totalRow.appendTableCell(''); // Sel kosong untuk kolom kuantitas
    totalRow.appendTableCell(''); // Sel kosong untuk kolom harga satuan
    totalRow.appendTableCell(dataInvoice.reduce(function(sum, item) {
      return sum + (item.kuantitas * item.harga);
    }, 0).toLocaleString('id-ID')).setText(dataInvoice.reduce(function(sum, item) {
      return sum + (item.kuantitas * item.harga);
    }, 0).toLocaleString('id-ID'));

    // Gabungkan tiga sel pertama di baris total
    totalRow.getCell(1).merge(); // Gabungkan sel kedua dengan sel pertama
    totalRow.getCell(2).merge(); // Gabungkan sel ketiga dengan sel kedua (yang sudah bergabung)

  } else {
    Logger.log('Tabel invoice tidak ditemukan dalam dokumen.');
  }
}

// Contoh penggunaan:
function main() {
  var templateId = '1g5mFGbpsTG2m0pzrfrxIZGlGXWt6xcVsKFMjzeYtL-Y'; // Ganti dengan ID file template Anda
  var dataInvoice = [
    { nama: "Jasa Desain", kuantitas: 1, harga: 500000 },
    { nama: "Cetak Brosur", kuantitas: 100, harga: 5000 },
    { nama: "Konsultasi", kuantitas: 2, harga: 250000 }
  ];
  var namaFileBaru = 'Invoice PT Maju Jaya - ' + Utilities.formatDate(new Date(), Session.getTimeZone(), 'yyyyMMdd');

  buatInvoiceDariTemplate(templateId, dataInvoice, namaFileBaru);
}
