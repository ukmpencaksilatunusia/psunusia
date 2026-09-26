// ==============================================================
// KONFIGURASI BACKEND UKM PENCAK SILAT UNUSIA
// KOMPATIBEL DENGAN FILE: code.html
// ==============================================================

// ⚠️ PASTE ID FOLDER GOOGLE DRIVE ANDA DI DALAM TANDA PETIK DI BAWAH INI
const FOLDER_DRIVE_ID = "1vsb6KHYtJY3g26P3lPV-IFlriQ5q-RPw";

// ==============================================================
// FUNGSI CEK KUOTA (Dibutuhkan untuk Polling Live di HTML)
// ==============================================================
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Hitung total pendaftar yang sudah masuk (dikurangi 1 baris Header)
  const totalPendaftar = Math.max(0, sheet.getLastRow() - 1);
  
  return ContentService
    .createTextOutput(JSON.stringify({ totalPendaftar: totalPendaftar }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
// FUNGSI PROSES PENDAFTARAN & UPLOAD
// ==============================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // Mencegah bentrokan jika ada yang submit bersamaan

  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const timestamp = new Date();

    // 1. Hitung urutan pendaftar (Baris 1 header, baris 2 pendaftar ke-1)
    const totalRow = sheet.getLastRow();
    const urutanPendaftar = totalRow; // Jika lastRow = 1, pendaftar ini ke-1
    const isPromoGratis = urutanPendaftar <= 10;
    const statusPendaftar = isPromoGratis 
      ? `Gratis (Pendaftar Tercepat Ke-${urutanPendaftar})` 
      : `Reguler (Pendaftar Ke-${urutanPendaftar})`;

    // 2. Akses folder penyimpanan di Drive
    const folder = DriveApp.getFolderById(FOLDER_DRIVE_ID);

    // 3. Simpan File Bukti Pembayaran (jika diunggah)
    let urlFileBayar = "Tidak Ada / Promo Gratis";
    if (data.file_payment && data.file_payment.data) {
      const blobBayar = Utilities.newBlob(
        Utilities.base64Decode(data.file_payment.data),
        data.file_payment.mimeType,
        `Payment_${data.nim}_${data.nama_lengkap}_${data.file_payment.filename}`
      );
      const fileBayar = folder.createFile(blobBayar);
      fileBayar.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFileBayar = fileBayar.getUrl();
    }

    // 4. Simpan File Bukti Follow Instagram
    let urlFileIG = "-";
    if (data.file_ig && data.file_ig.data) {
      const blobIG = Utilities.newBlob(
        Utilities.base64Decode(data.file_ig.data),
        data.file_ig.mimeType,
        `BuktiIG_${data.nim}_${data.nama_lengkap}_${data.file_ig.filename}`
      );
      const fileIG = folder.createFile(blobIG);
      fileIG.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFileIG = fileIG.getUrl();
    }

    // 5. Simpan File Bukti Follow TikTok
    let urlFileTikTok = "-";
    if (data.file_tiktok && data.file_tiktok.data) {
      const blobTikTok = Utilities.newBlob(
        Utilities.base64Decode(data.file_tiktok.data),
        data.file_tiktok.mimeType,
        `BuktiTikTok_${data.nim}_${data.nama_lengkap}_${data.file_tiktok.filename}`
      );
      const fileTikTok = folder.createFile(blobTikTok);
      fileTikTok.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFileTikTok = fileTikTok.getUrl();
    }

    // 6. Simpan Data ke Google Sheet Sesuai Urutan Kolom (A - R)
    sheet.appendRow([
      timestamp,                           // Col A: Timestamp
      data.nama_lengkap,                   // Col B: Nama Lengkap
      data.email,                          // Col C: Email
      "'" + data.nim,                      // Col D: NIM (dengan ' agar tidak dianggap formula)
      "'" + data.whatsapp,                 // Col E: No WhatsApp
      data.fakultas,                       // Col F: Fakultas
      data.prodi,                          // Col G: Program Studi
      data.kampus,                         // Col H: Pilihan Kampus
      data.unit_perguruan,                 // Col I: Unit Perguruan / Peminatan
      urlFileBayar,                        // Col J: Bukti Pembayaran (Link Drive)
      urlFileIG,                           // Col K: Bukti Follow IG (Link Drive)
      urlFileTikTok,                       // Col L: Bukti Follow TikTok (Link Drive)
      statusPendaftar,                     // Col M: Status Promo / Pendaftar
      data.timestamp || timestamp.toString() // Col N: Timestamp dari Browser (untuk backup)
    ]);

    // 7. Kirim Email Notifikasi Konfirmasi
    kirimEmailKonfirmasi(data, statusPendaftar, isPromoGratis, urutanPendaftar);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "success", 
        message: "Pendaftaran berhasil!", 
        urutan: urutanPendaftar,
        status: statusPendaftar
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "error", 
        message: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================
// FUNGSI EMAIL NOTIFIKASI BERGAYA MODERN (HTML EMAIL)
// ==============================================================
function kirimEmailKonfirmasi(data, statusPendaftar, isPromoGratis, urutan) {
  const subjek = `[Konfirmasi Pendaftaran] UKM Pencak Silat UNUSIA - ${data.nama_lengkap}`;
  
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; background-color: #ffffff; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);">
      
      <!-- Header Email -->
      <div style="background: linear-gradient(135deg, #065f46 0%, #059669 100%); padding: 32px 24px; text-align: center; color: #ffffff;">
        <div style="background-color: rgba(255,255,255,0.2); width: 60px; height: 60px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 12px; font-weight: bold; font-size: 20px;">
          🥋
        </div>
        <h1 style="margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 0.5px;">UKM PENCAK SILAT UNUSIA</h1>
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.9;">Universitas Nahdlatul Ulama Indonesia</p>
      </div>

      <!-- Isi Email -->
      <div style="padding: 28px 24px; color: #1f2937;">
        <h2 style="color: #065f46; font-size: 18px; margin-top: 0; font-weight: 600;">Halo, ${data.nama_lengkap}! 👋</h2>
        <p style="line-height: 1.6; font-size: 14px; color: #4b5563; margin-bottom: 20px;">
          Selamat! Formulir pendaftaran Anda untuk menjadi bagian dari keluarga besar <strong>UKM Pencak Silat UNUSIA</strong> telah berhasil kami terima.
        </p>

        <!-- Status Card Promo / Regular -->
        <div style="background-color: ${isPromoGratis ? '#ecfdf5' : '#fffbe2'}; border: 1px solid ${isPromoGratis ? '#a7f3d0' : '#fef08a'}; border-left: 5px solid ${isPromoGratis ? '#10b981' : '#f59e0b'}; padding: 16px; border-radius: 10px; margin: 20px 0;">
          <p style="margin: 0; font-size: 14px; font-weight: 700; color: ${isPromoGratis ? '#065f46' : '#92400e'};">
            🎉 Pendaftar Ke-${urutan}
          </p>
          <p style="margin: 4px 0 0 0; font-size: 13px; color: ${isPromoGratis ? '#047857' : '#b45309'};">
            Status Pendaftaran: <strong>${statusPendaftar}</strong>
          </p>
        </div>

        <!-- Tabel Rincian Data -->
        <h3 style="font-size: 15px; color: #111827; margin-top: 24px; margin-bottom: 12px; border-bottom: 2px solid #f3f4f6; padding-bottom: 6px; font-weight: 600;">Ringkasan Data Anda</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; color: #374151;">
          <tr>
            <td style="padding: 8px 0; font-weight: 600; width: 35%; color: #6b7280;">NIM</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.nim}</td>
          </tr>
          <tr style="background-color: #fafafa;">
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">Email</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.email}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">Fakultas</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.fakultas}</td>
          </tr>
          <tr style="background-color: #fafafa;">
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">Program Studi</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.prodi}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">Kampus</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.kampus || "-"}</td>
          </tr>
          <tr style="background-color: #fafafa;">
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">Unit Perguruan / Peminatan</td>
            <td style="padding: 8px 0; font-weight: 500;">: <strong>${data.unit_perguruan || "-"}</strong></td>
          </tr>
          <tr>
            <td style="padding: 8px 0; font-weight: 600; color: #6b7280;">No. WhatsApp</td>
            <td style="padding: 8px 0; font-weight: 500;">: ${data.whatsapp}</td>
          </tr>
        </table>

        <!-- Info Instruksi Selanjutnya -->
        <div style="margin-top: 24px; padding: 16px; background-color: #f8fafc; border-radius: 10px; font-size: 13px; color: #475569; line-height: 1.5; border: 1px dashed #cbd5e1;">
          <strong style="color: #0f172a;">📌 Langkah Selanjutnya:</strong><br>
          Tim pengurus akan verifikasi berkas Anda. Selanjutnya, Anda akan dihubungi atau dimasukkan ke dalam <strong>Group WhatsApp Calon Anggota Baru</strong>. Pastikan nomor WhatsApp Anda tetap aktif.
        </div>

        <br>
        <p style="font-size: 13px; color: #4b5563; margin-bottom: 0;">Salam hangat & lestari,<br><strong style="color: #065f46;">Pengurus UKM Pencak Silat UNUSIA</strong></p>
      </div>

      <!-- Footer Email -->
      <div style="background-color: #f9fafb; padding: 16px; text-align: center; font-size: 11px; color: #9ca3af; border-top: 1px solid #f3f4f6;">
        Email ini dikirim otomatis oleh Sistem Pendaftaran UKM Pencak Silat UNUSIA.<br>
        Kunjungi IG: <a href="https://instagram.com/ukm.pencaksilatunusia" style="color: #059669; text-decoration: none; font-weight: 600;">@ukm.pencaksilatunusia</a>
      </div>

    </div>
  `;

  // Mengirimkan Email
  MailApp.sendEmail({
    to: data.email,
    subject: subjek,
    htmlBody: htmlBody
  });
}
