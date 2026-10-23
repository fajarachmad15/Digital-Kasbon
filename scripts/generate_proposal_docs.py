"""
Script to generate Proposal Documents in PDF, DOCX, HTML, and Markdown
for Centralized Corporate Email & Digital Kasbon System.
"""
import os
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"D:\Project Aplikasi\Digital-Kasbon"

HTML_CONTENT = """<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Proposal Permohonan Email Corporate Terpusat - Digital Kasbon</title>
  <style>
    @page {
      size: A4;
      margin: 20mm 15mm 20mm 15mm;
    }
    * {
      box-sizing: border-box;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    body {
      font-family: 'Segoe UI', Arial, Helvetica, sans-serif;
      color: #1e293b;
      line-height: 1.5;
      background-color: #f8fafc;
      margin: 0;
      padding: 24px;
    }
    .page-container {
      max-width: 860px;
      margin: 0 auto;
      background: #ffffff;
      padding: 40px 48px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.08);
      border-radius: 8px;
    }
    .no-print-bar {
      max-width: 860px;
      margin: 0 auto 20px auto;
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: #1e293b;
      color: #fff;
      padding: 12px 20px;
      border-radius: 8px;
    }
    .btn-print {
      background: #2563eb;
      color: white;
      border: none;
      padding: 8px 18px;
      border-radius: 6px;
      font-weight: 600;
      font-size: 14px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }
    .btn-print:hover {
      background: #1d4ed8;
    }
    .header-memo {
      border-bottom: 3px double #0f172a;
      padding-bottom: 16px;
      margin-bottom: 24px;
    }
    .company-title {
      font-size: 20px;
      font-weight: 800;
      color: #0f172a;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      margin: 0 0 4px 0;
    }
    .doc-subtitle {
      font-size: 13px;
      color: #64748b;
      margin: 0 0 16px 0;
      font-weight: 500;
    }
    .meta-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 20px;
      font-size: 13px;
    }
    .meta-table td {
      padding: 6px 10px;
      vertical-align: top;
    }
    .meta-table td.label {
      width: 130px;
      font-weight: 700;
      color: #334155;
      background: #f1f5f9;
      border: 1px solid #e2e8f0;
    }
    .meta-table td.val {
      border: 1px solid #e2e8f0;
      color: #0f172a;
    }
    h2 {
      font-size: 16px;
      color: #0f172a;
      border-left: 4px solid #2563eb;
      padding-left: 10px;
      margin: 28px 0 12px 0;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }
    p, li {
      font-size: 13.5px;
      color: #334155;
      text-align: justify;
    }
    ul, ol {
      padding-left: 20px;
      margin: 8px 0;
    }
    li {
      margin-bottom: 6px;
    }
    .comparison-table {
      width: 100%;
      border-collapse: collapse;
      margin: 18px 0;
      font-size: 12px;
    }
    .comparison-table th {
      background: #0f172a;
      color: #ffffff;
      padding: 10px 12px;
      font-weight: 600;
      text-align: left;
      border: 1px solid #0f172a;
    }
    .comparison-table td {
      padding: 8px 10px;
      border: 1px solid #cbd5e1;
      vertical-align: top;
      line-height: 1.4;
    }
    .comparison-table tr:nth-child(even) {
      background-color: #f8fafc;
    }
    .badge-bad {
      display: inline-block;
      background: #fee2e2;
      color: #991b1b;
      padding: 2px 6px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 11px;
      margin-bottom: 4px;
    }
    .badge-good {
      display: inline-block;
      background: #dcfce7;
      color: #166534;
      padding: 2px 6px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 11px;
      margin-bottom: 4px;
    }
    .box-quote {
      background: #eff6ff;
      border-left: 4px solid #3b82f6;
      padding: 12px 16px;
      border-radius: 4px;
      margin: 16px 0;
      font-size: 13px;
    }
    .box-role {
      background: #f1f5f9;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 14px 18px;
      margin: 14px 0;
      font-size: 13px;
    }
    .box-role ul {
      margin: 6px 0 0 0;
      padding-left: 18px;
    }
    .signature-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 20px;
      margin-top: 40px;
      page-break-inside: avoid;
    }
    .sig-card {
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 14px;
      text-align: center;
      background: #fafafa;
    }
    .sig-role {
      font-size: 12px;
      font-weight: 700;
      color: #475569;
      margin-bottom: 60px;
      text-transform: uppercase;
    }
    .sig-name {
      font-size: 13px;
      font-weight: 700;
      color: #0f172a;
      border-top: 1px solid #0f172a;
      padding-top: 6px;
      display: inline-block;
      width: 85%;
    }
    .sig-title {
      font-size: 11px;
      color: #64748b;
      margin-top: 2px;
    }
    @media print {
      body {
        background: #fff;
        padding: 0;
      }
      .no-print-bar {
        display: none !important;
      }
      .page-container {
        box-shadow: none;
        padding: 0;
        max-width: 100%;
      }
      .page-break {
        page-break-before: always;
      }
    }
  </style>
</head>
<body>

  <div class="no-print-bar">
    <div><strong>📄 Format Cetak Dokumen Resmi</strong> (Klik tombol di samping untuk mencetak / download PDF)</div>
    <button class="btn-print" onclick="window.print()">🖨️ Cetak / Simpan sebagai PDF</button>
  </div>

  <div class="page-container">
    <div class="header-memo">
      <div class="company-title">KAWAN LAMA GROUP</div>
      <div class="doc-subtitle">INTERNAL MEMORANDUM &amp; JUSTIFIKASI TEKNIS SISTEM DIGITALISASI KASBON</div>
      <table class="meta-table">
        <tr>
          <td class="label">Nomor Dokumen</td>
          <td class="val"><strong>012/MEMO-IT-FIN/DKP/X/2026</strong></td>
          <td class="label">Tanggal</td>
          <td class="val">23 Oktober 2026</td>
        </tr>
        <tr>
          <td class="label">Kepada</td>
          <td class="val"><strong>Head of Corporate IT &amp; Infrastructure</strong><br><strong>General Manager Finance &amp; Accounting</strong></td>
          <td class="label">Dari</td>
          <td class="val">Tim Inisiatif Digitalisasi Kasbon (Petty Cash Modernization)</td>
        </tr>
        <tr>
          <td class="label">Perihal</td>
          <td class="val" colspan="3"><strong>Permohonan Pengadaan &amp; Aktivasi 1 Akun Email Corporate Terpusat untuk Sentralisasi Sistem Digital Kasbon (Enterprise Petty Cash)</strong></td>
        </tr>
        <tr>
          <td class="label">Sifat &amp; Klasifikasi</td>
          <td class="val" colspan="3"><span style="color:#b91c1c; font-weight:700;">PENTING &amp; SEGERA</span> (Efisiensi Operasional, Keamanan Data &amp; Tata Kelola Keuangan)</td>
        </tr>
      </table>
    </div>

    <h2>1. Latar Belakang Masalah (Sistem Lama 58 Langkah)</h2>
    <p>
      Berdasarkan dokumentasi teknis sistem lama (<em>Bon Sementara Spreadsheet &amp; Google Form</em>), proses pendaftaran dan operasional kas kecil di setiap toko/store memiliki beban teknis manual yang sangat berat dan berisiko tinggi terhadap keamanan data korporasi.
    </p>
    <p>
      Untuk menambahkan <strong>1 store baru</strong> saja, teknisi/admin harus menjalankan <strong>58 tahapan manual (total 203 langkah teknis)</strong>, meliputi:
    </p>
    <ul>
      <li>Membuat dan mengelola akun Gmail personal terpisah per store (contoh: <code>bon.sementara501@gmail.com</code>).</li>
      <li>Menginstal add-on pihak ketiga (<em>Mail Merge</em>) dengan kuota trial gratis yang terbatas hanya 30 baris email/akun, sehingga sering macet di tengah jalan.</li>
      <li>Duplikasi manual 5 Google Form (Pengajuan, Disetujui, Diberikan, Diterima, Realisasi) dan 1 Google Spreadsheet per store.</li>
      <li>Transfer kepemilikan file bolak-balik antara akun alter dan akun store, serta menyetel proteksi rentang satu demi satu (20+ kali setel izin sheet).</li>
      <li><em>Copy-paste</em> manual daftar email dan sintaks ekspresi reguler (Regex) NIP/NIK karyawan ke dalam form.</li>
      <li>Membuka hak akses Google Drive ke publik (<em>"Siapa saja yang memiliki link dapat melihat"</em>).</li>
    </ul>

    <h2>2. Solusi Usulan: Sentralisasi 1 Akun Corporate</h2>
    <p>
      Seluruh sistem kasbon kini telah dimodernisasi menjadi <strong>Web App Suite Terpadu (Google Apps Script + Supabase Cloud)</strong>. Seluruh proses di atas dapat dihilangkan secara permanen hanya dengan menyediakan <strong>1 (satu) Akun Google Workspace Corporate resmi</strong> (misalnya: <code>kasbon.digital@kawanlamacorp.com</code>).
    </p>
    <div class="box-quote">
      <strong>Hasil Sentralisasi:</strong> Penambahan toko baru kini <strong>TIDAK PERLU</strong> lagi membuat akun email baru, tidak perlu membuat form baru, dan tidak perlu menyalin sheet. Cukup 10 detik memasukkan Kode Store di Portal Admin, sistem langsung aktif otomatis!
    </div>

    <div class="page-break"></div>

    <h2>3. Matriks Komparasi Mendalam: BEFORE vs AFTER</h2>
    <table class="comparison-table">
      <thead>
        <tr>
          <th style="width: 18%;">Parameter</th>
          <th style="width: 38%;">Sistem Lama (Email Terpisah per Store)</th>
          <th style="width: 44%;">Sistem Baru (1 Email Corporate Terpusat)</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><strong>Jumlah Akun Email</strong></td>
          <td>
            <span class="badge-bad">Rentan &amp; Tercecer</span><br>
            Puluhan akun Gmail gratisan personal (<code>bon.sementaraXXX@gmail.com</code>) + akun alter. Rawan lupa password &amp; OTP.
          </td>
          <td>
            <span class="badge-good">Terkendali &amp; Terpusat</span><br>
            <strong>Tepat 1 Akun Resmi Google Workspace</strong> (misal: <code>kasbon.digital@kawanlamacorp.com</code>) di bawah kendali penuh IT HO.
          </td>
        </tr>
        <tr>
          <td><strong>Waktu Setup Store Baru</strong></td>
          <td>
            <span class="badge-bad">Sangat Lambat</span><br>
            <strong>2 - 3 Jam per Store:</strong> Menjalankan 58 halaman SOP tutorial (203 langkah manual kloning file &amp; setting izin).
          </td>
          <td>
            <span class="badge-good">Instan (Efisiensi 99.8%)</span><br>
            <strong>10 Detik:</strong> Cukup daftarkan Kode &amp; Nama Store di Portal Admin. Form otomatis melayani toko baru seketika.
          </td>
        </tr>
        <tr>
          <td><strong>Add-On &amp; Biaya Pihak Ketiga</strong></td>
          <td>
            <span class="badge-bad">Ketergantungan Eksternal</span><br>
            Wajib install add-on pihak ketiga (<em>Mail Merge</em>). Dibatasi kuota gratis 30 baris/akun; risiko biaya subscription berulang.
          </td>
          <td>
            <span class="badge-good">Bebas Biaya (Rp 0,-)</span><br>
            <strong>100% Native Google Apps Script:</strong> Menggunakan library internal tanpa add-on eksternal, bebas risiko limit trial.
          </td>
        </tr>
        <tr>
          <td><strong>Kapasitas Pengiriman Email</strong></td>
          <td>
            <span class="badge-bad">Sering Terhenti</span><br>
            Dibatasi kuota Gmail personal (maks 100 email/hari). Sering gagal kirim saat volume kasbon akhir bulan memuncak.
          </td>
          <td>
            <span class="badge-good">Reliabel &amp; Stabil</span><br>
            <strong>1.500 - 2.000 email/hari</strong> (kuota resmi Google Workspace Enterprise). Notifikasi approval langsung sampai detik itu juga.
          </td>
        </tr>
        <tr>
          <td><strong>Penyimpanan Bukti (Drive)</strong></td>
          <td>
            <span class="badge-bad">Terfragmentasi</span><br>
            Bukti nota tercecer di puluhan drive akun personal toko yang tidak terawasi oleh Head Office.
          </td>
          <td>
            <span class="badge-good">Satu Pintu Korporat</span><br>
            Tersimpan rapi di Google Drive akun corporate dengan folder otomatis per bulan (<code>Digital Kasbon Bukti/YYYY-MM/No_Kasbon</code>).
          </td>
        </tr>
        <tr>
          <td><strong>Keamanan &amp; Kepatuhan Data</strong></td>
          <td>
            <span class="badge-bad">Risiko Kebocoran Tinggi</span><br>
            Sheet dan folder diatur ke publik (<em>"Siapa saja yang memiliki link dapat melihat"</em>). Sangat berbahaya bagi data keuangan.
          </td>
          <td>
            <span class="badge-good">Enterprise Grade Security</span><br>
            Terkunci aman di dalam domain internal <code>@kawanlamacorp.com</code>, terenkripsi SSL, dan tersimpan di database aman Supabase.
          </td>
        </tr>
        <tr>
          <td><strong>Otorisasi Pemohon (SSO)</strong></td>
          <td>
            <span class="badge-bad">Bisa Dipalsukan</span><br>
            Nama dan email diketik bebas oleh pemohon di form tanpa verifikasi identitas akun yang sesungguhnya.
          </td>
          <td>
            <span class="badge-good">Terkunci Akun Resmi</span><br>
            <strong>Auto-Detect Email Google SSO:</strong> Email terisi otomatis sesuai akun login dan <strong>dikunci permanen (readonly)</strong>.
          </td>
        </tr>
        <tr>
          <td><strong>Pencegahan Fraud (Kasbon Ganda)</strong></td>
          <td>
            <span class="badge-bad">Human Error Kasir</span><br>
            Hanya mengandalkan ketelitian kasir mengecek baris sheet; kasbon ganda berulang kali lolos pencairan.
          </td>
          <td>
            <span class="badge-good">Anti-Double Kasbon Guard</span><br>
            Sistem otomatis memblokir pengajuan jika email atau NIP masih memiliki kasbon aktif yang belum diselesaikan.
          </td>
        </tr>
        <tr>
          <td><strong>Rekonsiliasi &amp; Audit Finance</strong></td>
          <td>
            <span class="badge-bad">Tutup Buku Lambat</span><br>
            Finance HO harus meminta link puluhan spreadsheet toko yang formatnya sering rusak terhapus oleh user.
          </td>
          <td>
            <span class="badge-good">Real-Time Consolidation</span><br>
            <strong>1 Dashboard Monitoring Nasional:</strong> Rekapitulasi transaksi seluruh toko terpantau live dalam satu tampilan.
          </td>
        </tr>
      </tbody>
    </table>

    <h2>4. Fungsi Khusus Akun Corporate yang Dimohonkan</h2>
    <div class="box-role">
      <strong>Akun: <code>kasbon.digital@kawanlamacorp.com</code> (Akun Sistem / Non-Personal)</strong>
      <ul>
        <li><strong>Host Backend Script:</strong> Menjalankan 6 modul aplikasi Google Apps Script terpusat.</li>
        <li><strong>Automated Mail Dispatcher:</strong> Mengirim notifikasi status kasbon, link persetujuan Manager, dan konfirmasi pencairan Kasir.</li>
        <li><strong>Central Drive Repository:</strong> Menampung folder arsip dokumen dan nota struk petty cash dari seluruh unit bisnis.</li>
      </ul>
    </div>

    <h2>5. Rekomendasi &amp; Permohonan Persetujuan</h2>
    <p>
      Mengingat manfaat efisiensi waktu kerja sebesar <strong>99.8%</strong>, penghapusan potensi biaya add-on berbayar, serta peningkatan standar keamanan data keuangan perusahaan, kami memohon kesediaan Bapak/Ibu Manajemen IT &amp; Finance untuk mengesahkan dan menerbitkan 1 akun corporate tersebut.
    </p>

    <div class="signature-grid">
      <div class="sig-card">
        <div class="sig-role">Diajukan Oleh:</div>
        <div class="sig-name">Achmad Fajar Setiawan</div>
        <div class="sig-title">Project Leader Digital Kasbon</div>
      </div>
      <div class="sig-card">
        <div class="sig-role">Disetujui Oleh (IT):</div>
        <div class="sig-name">Head of Corporate IT</div>
        <div class="sig-title">IT Infrastructure &amp; Security</div>
      </div>
      <div class="sig-card">
        <div class="sig-role">Disetujui Oleh (Finance):</div>
        <div class="sig-name">GM Finance &amp; Accounting</div>
        <div class="sig-title">Finance &amp; Corporate Control</div>
      </div>
    </div>
  </div>

</body>
</html>
"""

def generate_html():
    out_html = os.path.join(BASE_DIR, "Proposal_Email_Corporate_Centralisasi_Kasbon.html")
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(HTML_CONTENT)
    print(f"[OK] HTML generated: {out_html}")
    return out_html

def generate_pdf(html_path):
    out_pdf = os.path.join(BASE_DIR, "Proposal_Email_Corporate_Centralisasi_Kasbon.pdf")
    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    cmd = [
        "powershell",
        "-Command",
        f'Start-Process -FilePath "{edge_exe}" -ArgumentList "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=`"{out_pdf}`"", "`"{html_path}`"" -Wait'
    ]
    subprocess.run(cmd, check=True)
    if os.path.exists(out_pdf):
        size_kb = os.path.getsize(out_pdf) / 1024
        print(f"[OK] PDF generated: {out_pdf} ({size_kb:.1f} KB)")
    else:
        print("[ERR] PDF generation failed")
    return out_pdf

def generate_docx():
    out_docx = os.path.join(BASE_DIR, "Proposal_Email_Corporate_Centralisasi_Kasbon.docx")
    doc = Document()

    # Page setup (A4, 2cm margins)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Document Header
    p_comp = doc.add_paragraph()
    r_comp = p_comp.add_run("KAWAN LAMA GROUP")
    r_comp.bold = True
    r_comp.font.size = Pt(16)
    r_comp.font.color.rgb = RGBColor(15, 23, 42)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("INTERNAL MEMORANDUM & JUSTIFIKASI TEKNIS SISTEM DIGITALISASI KASBON")
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # Meta Table
    meta_table = doc.add_table(rows=4, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Nomor Dokumen", "012/MEMO-IT-FIN/DKP/X/2026", "Tanggal", "23 Oktober 2026"),
        ("Kepada", "Head of Corporate IT & Infrastructure\nGM Finance & Accounting", "Dari", "Tim Inisiatif Digitalisasi Kasbon"),
        ("Perihal", "Permohonan Pengadaan & Aktivasi 1 Akun Email Corporate Terpusat untuk Sentralisasi Sistem Digital Kasbon", "", ""),
        ("Sifat & Status", "PENTING & SEGERA (Efisiensi Operasional, Keamanan Data & Tata Kelola Keuangan)", "", "")
    ]

    for i, row in enumerate(meta_table.rows):
        d = meta_data[i]
        if i in (2, 3):
            # merged row
            row.cells[0].text = d[0]
            row.cells[1].text = d[1]
            row.cells[1].merge(row.cells[2]).merge(row.cells[3])
            row.cells[0].paragraphs[0].runs[0].bold = True
        else:
            row.cells[0].text = d[0]
            row.cells[1].text = d[1]
            row.cells[2].text = d[2]
            row.cells[3].text = d[3]
            row.cells[0].paragraphs[0].runs[0].bold = True
            row.cells[2].paragraphs[0].runs[0].bold = True

    # Section 1
    doc.add_heading("1. Latar Belakang Masalah (Sistem Lama 58 Langkah)", level=2)
    p = doc.add_paragraph(
        "Berdasarkan tinjauan operasional sistem lama (Bon Sementara Spreadsheet & Google Form), proses pendaftaran "
        "dan operasional kas kecil di setiap toko/store memiliki beban teknis manual yang sangat berat dan berisiko tinggi terhadap keamanan data korporasi. "
        "Untuk mendaftarkan 1 store baru saja, tim wajib menjalankan 58 tahapan manual (total 203 langkah teknis individual), yang meliputi:"
    )
    doc.add_paragraph("• Membuat dan mengelola akun Gmail personal terpisah per store (contoh: bon.sementara501@gmail.com).")
    doc.add_paragraph("• Menginstal add-on pihak ketiga (Mail Merge) dengan kuota trial terbatas hanya 30 baris email/akun.")
    doc.add_paragraph("• Duplikasi manual 5 Google Form dan 1 Google Spreadsheet per store.")
    doc.add_paragraph("• Transfer kepemilikan file bolak-balik antara akun alter dan akun store, serta menyetel proteksi rentang satu per satu.")
    doc.add_paragraph("• Copy-paste manual daftar email dan sintaks regex NIP/NIK karyawan ke dalam form.")
    doc.add_paragraph("• Membuka hak akses Google Drive ke publik ('Siapa saja yang memiliki link dapat melihat').")

    # Section 2
    doc.add_heading("2. Solusi Usulan: Sentralisasi 1 Akun Corporate", level=2)
    doc.add_paragraph(
        "Seluruh sistem kasbon kini telah dimodernisasi menjadi Web App Suite Terpadu (Google Apps Script + Supabase Cloud). "
        "Seluruh proses di atas dapat dihilangkan secara permanen hanya dengan menyediakan 1 (satu) Akun Google Workspace Corporate resmi "
        "(misalnya: kasbon.digital@kawanlamacorp.com)."
    )
    doc.add_paragraph(
        "Hasil Sentralisasi: Penambahan toko baru kini TIDAK PERLU lagi membuat akun email baru, tidak perlu membuat form baru, "
        "dan tidak perlu menyalin sheet. Cukup 10 detik memasukkan Kode Store di Portal Admin, sistem langsung aktif otomatis."
    )

    # Section 3: Before vs After Table
    doc.add_heading("3. Matriks Komparasi Mendalam: BEFORE vs AFTER", level=2)
    comp_table = doc.add_table(rows=1, cols=3)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = comp_table.rows[0].cells
    hdr[0].text = "Parameter Evaluasi"
    hdr[1].text = "Sistem Lama (Email Terpisah)"
    hdr[2].text = "Sistem Baru (1 Email Corporate)"
    for cell in hdr:
        cell.paragraphs[0].runs[0].bold = True

    comp_rows = [
        ("Jumlah Akun Email", "Puluhan akun Gmail gratisan (bon.sementaraXXX@gmail.com). Rawan lupa password & OTP.", "Tepat 1 Akun Resmi Google Workspace (kasbon.digital@kawanlamacorp.com) di bawah IT."),
        ("Waktu Setup Toko Baru", "2 - 3 Jam per Store: Menjalankan 58 halaman SOP tutorial (203 langkah manual).", "10 Detik (Efisiensi 99.8%): Cukup input Kode Store di Portal Admin satu kali."),
        ("Add-On & Biaya Eksternal", "Wajib install add-on pihak ketiga (Mail Merge) dengan limit kuota gratis 30 baris/akun.", "100% Native Google Apps Script: Tanpa add-on berbayar eksternal (Rp 0,- tambahan)."),
        ("Kapasitas Kirim Email", "Maksimal 100 email/hari (Gmail gratis). Sering macet saat penutupan kasbon.", "1.500 - 2.000 email/hari (Google Workspace). Notifikasi real-time tanpa batas."),
        ("Penyimpanan File Bukti", "Tersebar di puluhan drive personal masing-masing toko, tanpa kontrol pusat.", "1 Master Google Drive Folder korporat (Digital Kasbon Bukti / YYYY-MM / No_Kasbon)."),
        ("Keamanan & Kepatuhan Data", "Link disetel publik ('Anyone with link can view'), rawan kebocoran data kas.", "Enterprise RBAC: Terkunci di domain internal @kawanlamacorp.com, database SSL."),
        ("Otorisasi Pemohon (SSO)", "Nama & email diketik manual bebas tanpa validasi, rawan pemalsuan identitas.", "Auto-Detect Google SSO: Email terisi otomatis sesuai akun login & dikunci permanen."),
        ("Pencegahan Fraud Kasbon", "Mengandalkan ketelitian mata kasir; sering terjadi double kasbon sebelum setor nota.", "Anti-Double Kasbon Guard: Otomatis memblokir pengajuan jika kasbon lama belum tuntas."),
        ("Rekonsiliasi & Audit Finance", "Finance HO harus mengumpulkan puluhan link sheet toko yang sering rusak rumusnya.", "1 Layar Dashboard Monitoring Real-Time: Rekapitulasi nasional live dalam 1 klik.")
    ]

    for param, old_val, new_val in comp_rows:
        row = comp_table.add_row().cells
        row[0].text = param
        row[0].paragraphs[0].runs[0].bold = True
        row[1].text = old_val
        row[2].text = new_val

    # Section 4
    doc.add_heading("4. Rincian Akun Corporate yang Dimohonkan", level=2)
    doc.add_paragraph("• Nama Akun: kasbon.digital@kawanlamacorp.com (Akun Sistem / Non-Personal)")
    doc.add_paragraph("• Display Name: Digital Kasbon System - Kawan Lama Group")
    doc.add_paragraph("• Fungsi: Script Host, Dispatcher Notifikasi Email Otomatis, dan Master File Repository.")

    # Section 5
    doc.add_heading("5. Rekomendasi & Permohonan Persetujuan", level=2)
    doc.add_paragraph(
        "Mengingat efisiensi waktu kerja sebesar 99.8%, penghapusan biaya add-on, dan pengamanan aset data keuangan perusahaan, "
        "kami memohon kesediaan Manajemen IT dan Finance untuk mengesahkan pengadaan 1 akun corporate tersebut."
    )

    # Signatures
    doc.add_paragraph("\n")
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_headers = ["Diajukan Oleh:", "Disetujui Oleh (IT):", "Disetujui Oleh (Finance):"]
    sig_names = [
        "Achmad Fajar Setiawan\nProject Leader Digital Kasbon",
        "Head of Corporate IT\nIT Infrastructure & Security",
        "GM Finance & Accounting\nFinance & Corporate Control"
    ]
    for i in range(3):
        sig_table.rows[0].cells[i].text = sig_headers[i]
        sig_table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        sig_table.rows[1].cells[i].text = f"\n\n\n_______________________\n{sig_names[i]}"

    doc.save(out_docx)
    print(f"[OK] DOCX generated: {out_docx}")
    return out_docx

def generate_markdown_copy():
    src_md = os.path.join(r"C:\Users\Admin\.gemini\antigravity\brain\417d33a9-34dc-4204-a6a7-ce495039e834", "proposal_email_corporate_centralisasi_kasbon.md")
    dst_md = os.path.join(BASE_DIR, "Proposal_Email_Corporate_Centralisasi_Kasbon.md")
    if os.path.exists(src_md):
        with open(src_md, "r", encoding="utf-8") as f_in:
            content = f_in.read()
        with open(dst_md, "w", encoding="utf-8") as f_out:
            f_out.write(content)
        print(f"[OK] Markdown copied: {dst_md}")

def main():
    print("Generating Proposal Documents in multiple downloadable formats...")
    html_file = generate_html()
    pdf_file = generate_pdf(html_file)
    docx_file = generate_docx()
    generate_markdown_copy()
    print("ALL PROPOSAL FILES READY FOR DOWNLOAD & PRINTING!")

if __name__ == "__main__":
    main()
