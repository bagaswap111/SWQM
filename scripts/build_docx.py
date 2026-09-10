#!/usr/bin/env python3
"""
Script to generate PROPOSAL_SWQMS.docx from markdown content.
Run this script to create the DOCX file in the output directory.
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

def add_heading_with_emoji(doc, text, level):
    """Add heading with emoji support"""
    heading = doc.add_heading(level=level)
    run = heading.add_run(text)
    return heading

def add_table_with_data(doc, headers, data, caption=None):
    """Add a formatted table with headers and data"""
    if caption:
        doc.add_paragraph(caption, style='Caption')
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Header row
    header_row = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_row[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
    
    # Data rows
    for row_data in data:
        row = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            row[i].text = str(cell_data)
            for paragraph in row[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
    
    return table

def build_proposal():
    """Build the complete proposal document"""
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Title Page
    title = doc.add_heading('PROPOSAL TEKNIS & KOMERSIAL', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(20)
        run.font.bold = True
    
    subtitle = doc.add_heading('SISTEM MONITORING KUALITAS AIR BERBASIS IoT', level=2)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.font.size = Pt(16)
    
    smart_title = doc.add_paragraph('Smart Water Quality Monitoring System (SWQMS)')
    smart_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in smart_title.runs:
        run.font.italic = True
        run.font.size = Pt(14)
    
    doc.add_paragraph()
    doc.add_paragraph('_' * 60).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    # Submission info
    doc.add_paragraph('Diajukan Kepada:', style='Heading 3')
    doc.add_paragraph('[Nama Perusahaan / Klien]')
    doc.add_paragraph()
    
    doc.add_paragraph('Disusun Oleh:', style='Heading 3')
    doc.add_paragraph('[Tim Engineering IoT]')
    doc.add_paragraph()
    
    doc.add_paragraph(f'Tanggal: 10 September 2026')
    doc.add_paragraph(f'Nomor Proposal: SWQMS-2026-001')
    doc.add_paragraph(f'Versi: 1.0')
    
    # Table of Contents placeholder
    doc.add_page_break()
    doc.add_heading('DAFTAR ISI', level=1)
    doc.add_paragraph('(Daftar isi akan digenerate otomatis oleh Word)')
    doc.add_paragraph()
    
    # Section 1: Ringkasan Eksekutif
    doc.add_heading('1. RINGKASAN EKSEKUTIF', level=1)
    doc.add_paragraph(
        'Smart Water Quality Monitoring System (SWQMS) adalah solusi Internet of Things (IoT) '
        'terintegrasi untuk memantau kualitas air secara real-time pada kolam budidaya, reservoir, '
        'atau badan air lainnya. Sistem ini dirancang untuk beroperasi secara otonom 24/7 dengan '
        'sumber daya mandiri (solar panel + baterai), mampu bekerja di kedalaman 1–3 meter, dan '
        'terintegrasi langsung dengan sistem ERP perusahaan.'
    )
    
    doc.add_paragraph('Highlight Utama:', style='Heading 2')
    highlights = [
        '✅ 5 Parameter Kualitas Air: pH, EC/TDS/Salinitas, DO (Dissolved Oxygen), Suhu, dan monitoring kincir',
        '✅ Otonom Energi: Solar panel 20W + baterai 12V 7.2Ah (backup 48+ jam)',
        '✅ Dual Connectivity: WiFi (MQTT → ThingsBoard) + GSM (SMS Alert)',
        '✅ Integrasi ERP: REST API untuk sinkronisasi data otomatis',
        '✅ Smart Recommendation: AI-based saran perbaikan kualitas air & pakan',
        '✅ Industrial Grade: Sensor IP68, casing waterproof, floating design'
    ]
    for highlight in highlights:
        doc.add_paragraph(highlight, style='List Bullet')
    
    doc.add_paragraph(
        'Investasi Total: Rp 17.258.500 per unit dengan ROI (Return on Investment) diperkirakan '
        '6–9 bulan melalui efisiensi operasional dan pencegahan kerugian.',
        style='Intense Quote'
    )
    
    # Section 2: Latar Belakang
    doc.add_heading('2. LATAR BELAKANG', level=1)
    
    doc.add_heading('2.1 Permasalahan yang Dihadapi', level=2)
    doc.add_paragraph(
        'Dalam industri akuakultur dan pengelolaan sumber daya air, kualitas air merupakan faktor kritis '
        'yang menentukan kesehatan organisme air, efisiensi pakan, risiko kematian massal, dan kepatuhan regulasi.'
    )
    
    doc.add_paragraph('Tantangan saat ini:', style='Heading 3')
    challenges = [
        'Monitoring manual yang tidak kontinu → data tidak representatif',
        'Keterlambatan deteksi masalah → kerugian material besar',
        'Tidak ada integrasi dengan sistem ERP → data terisolasi',
        'Ketergantungan listrik PLN → rentan terhadap pemadaman',
        'Tidak ada sistem peringatan dini → reaksi lambat saat krisis'
    ]
    for challenge in challenges:
        doc.add_paragraph(challenge, style='List Number')
    
    doc.add_heading('2.2 Peluang Solusi', level=2)
    doc.add_paragraph(
        'Dengan perkembangan teknologi IoT, sensor industrial, dan cloud computing, kini memungkinkan '
        'untuk membangun sistem monitoring yang otonom, real-time, terintegrasi, dan cerdas.'
    )
    
    # Section 3: Tujuan Proyek
    doc.add_heading('3. TUJUAN PROYEK', level=1)
    
    doc.add_heading('3.1 Tujuan Umum', level=2)
    doc.add_paragraph(
        'Membangun sistem monitoring kualitas air berbasis IoT yang handal, otonom, dan terintegrasi '
        'untuk mendukung operasional budidaya/perusahaan secara efisien dan berkelanjutan.'
    )
    
    doc.add_heading('3.2 Tujuan Khusus', level=2)
    objectives = [
        'Memantau 5 parameter kualitas air secara real-time 24/7',
        'Memberikan peringatan dini via SMS & dashboard saat terjadi anomali',
        'Mendeteksi kegagalan peralatan (kincir mati) secara otomatis',
        'Menyediakan data historis untuk analisis tren dan audit',
        'Terintegrasi dengan sistem ERP perusahaan',
        'Memberikan rekomendasi cerdas untuk perbaikan kualitas air & manajemen pakan',
        'Mengurangi kerugian operasional minimal 30% dalam 1 tahun'
    ]
    for obj in objectives:
        doc.add_paragraph(obj, style='List Number')
    
    # Section 4: Manfaat & Value Proposition
    doc.add_heading('4. MANFAAT & VALUE PROPOSITION', level=1)
    
    doc.add_heading('4.1 Manfaat Operasional', level=2)
    table_headers = ['Aspek', 'Sebelum', 'Sesudah']
    table_data = [
        ['Monitoring', 'Manual, 1–2x/hari', 'Otomatis, setiap 1 menit'],
        ['Deteksi masalah', 'Reaktif (setelah kerugian)', 'Proaktif (peringatan dini)'],
        ['Data recording', 'Spreadsheet manual', 'Cloud database terstruktur'],
        ['Integrasi ERP', 'Tidak ada', 'REST API otomatis'],
        ['Respons time', 'Jam–hari', 'Menit (via SMS)']
    ]
    add_table_with_data(doc, table_headers, table_data, 'Tabel 4.1: Perbandingan Sebelum dan Sesudah')
    
    doc.add_heading('4.2 Manfaat Ekonomi', level=2)
    economic_benefits = [
        'Pengurangan kerugian akibat kematian ikan/udang: Rp 50–200 juta/tahun',
        'Efisiensi pakan melalui rekomendasi cerdas: 10–15% penghematan',
        'Pengurangan tenaga kerja monitoring manual: 1–2 orang',
        'ROI (Return on Investment): 6–9 bulan'
    ]
    for benefit in economic_benefits:
        doc.add_paragraph(benefit, style='List Bullet')
    
    # Section 5: Spesifikasi Teknis
    doc.add_heading('5. SPESIFIKASI TEKNIS SISTEM', level=1)
    
    doc.add_heading('5.1 Spesifikasi Sensor', level=2)
    sensor_headers = ['No', 'Parameter', 'Model', 'Range', 'Akurasi', 'Output', 'Proteksi']
    sensor_data = [
        ['1', 'pH', 'Industrial RS485', '0–14 pH', '±0.1 pH', 'Modbus RTU', 'IP68'],
        ['2', 'EC/TDS/Salinitas', 'Industrial RS485', '0–20000 µS/cm', '±1%', 'Modbus RTU', 'IP68'],
        ['3', 'Dissolved Oxygen', 'Optical RS485', '0–20 mg/L', '±0.1 mg/L', 'Modbus RTU', 'IP68'],
        ['4', 'Suhu', 'PT100 RS485', '-10–85°C', '±0.5°C', 'Modbus RTU', 'IP68'],
        ['5', 'Arus Kincir', 'SCT-013-030', '0–30A AC', '±1%', 'Analog 0–1V', 'Non-invasive']
    ]
    add_table_with_data(doc, sensor_headers, sensor_data, 'Tabel 5.1: Spesifikasi Sensor')
    
    doc.add_heading('5.2 Spesifikasi Mikrokontroler & Komunikasi', level=2)
    mcu_headers = ['Komponen', 'Spesifikasi']
    mcu_data = [
        ['MCU', 'ESP32 DevKit V1 (Dual-core 240MHz, WiFi 802.11 b/g/n, Bluetooth 4.2)'],
        ['Memory', '520KB SRAM, 4MB Flash'],
        ['Protokol', 'MQTT (ThingsBoard), HTTP/REST (ERP), AT Command (SMS)'],
        ['GSM Module', 'SIM800L Quad-band (850/900/1800/1900 MHz)'],
        ['RS485', 'MAX3485, 3.3V compatible, auto-direction'],
        ['Jarak Komunikasi', 'RS485: hingga 1200m, WiFi: 100m, GSM: nationwide']
    ]
    add_table_with_data(doc, mcu_headers, mcu_data, 'Tabel 5.2: Spesifikasi MCU')
    
    doc.add_heading('5.3 Spesifikasi Power Supply', level=2)
    power_headers = ['Komponen', 'Spesifikasi']
    power_data = [
        ['Solar Panel', 'Monocrystalline 20W, 12V, waterproof'],
        ['Charge Controller', 'PWM 10A, 12V/24V auto-detect, LCD display'],
        ['Battery', 'VRLA/SLA 12V 7.2Ah, maintenance-free'],
        ['Backup Time', '48+ jam tanpa sinar matahari'],
        ['Regulator', '2x LM2596 Buck Converter (5V 3A & 3.3V 2A)'],
        ['Proteksi', 'Fuse 5A, TVS Diode, overcharge/overdischarge protection']
    ]
    add_table_with_data(doc, power_headers, power_data, 'Tabel 5.3: Spesifikasi Power Supply')
    
    # Section 10: RAB
    doc.add_heading('10. RENCANA ANGGARAN BIAYA (RAB)', level=1)
    
    doc.add_heading('10.1 Bill of Materials (BOM)', level=2)
    
    doc.add_paragraph('A. Komponen Utama (Sensor & Mikrokontroler)', style='Heading 3')
    bom_a_headers = ['No', 'Komponen', 'Spesifikasi', 'Qty', 'Harga Satuan', 'Total']
    bom_a_data = [
        ['1', 'ESP32 DevKit V1', 'ESP32-WROOM-32, WiFi+BT', '1', 'Rp 75.000', 'Rp 75.000'],
        ['2', 'Sensor pH Industrial', 'RS485 Modbus, IP68, kabel 5m', '1', 'Rp 3.150.000', 'Rp 3.150.000'],
        ['3', 'Sensor EC/TDS/Salinity', 'RS485 Modbus, IP68, kabel 5m', '1', 'Rp 3.650.000', 'Rp 3.650.000'],
        ['4', 'Sensor DO Industrial', 'Optical RS485, IP68, kabel 5m', '1', 'Rp 8.500.000', 'Rp 8.500.000'],
        ['5', 'Sensor Suhu Industrial', 'PT100 RS485, IP68, kabel 5m', '1', 'Rp 450.000', 'Rp 450.000'],
        ['6', 'Modul RS485 to TTL', 'MAX3485, 3.3V compatible', '1', 'Rp 25.000', 'Rp 25.000'],
        ['7', 'SIM800L GSM Module', 'Quad-band, SMS+GPRS', '1', 'Rp 65.000', 'Rp 65.000'],
        ['8', 'SCT-013-030', 'Current Transformer 30A', '1', 'Rp 45.000', 'Rp 45.000'],
        ['', '', 'Subtotal A', '', '', 'Rp 16.011.000']
    ]
    add_table_with_data(doc, bom_a_headers, bom_a_data, 'Tabel 10.1: BOM Komponen Utama')
    
    doc.add_paragraph()
    doc.add_paragraph('B. Power Supply', style='Heading 3')
    bom_b_headers = ['No', 'Komponen', 'Spesifikasi', 'Qty', 'Harga Satuan', 'Total']
    bom_b_data = [
        ['10', 'Solar Panel', 'Monocrystalline 20W 12V', '1', 'Rp 220.000', 'Rp 220.000'],
        ['11', 'Solar Charge Controller', 'PWM 10A, LCD display', '1', 'Rp 110.000', 'Rp 110.000'],
        ['12', 'Baterai Kering', 'VRLA/SLA 12V 7.2Ah', '1', 'Rp 225.000', 'Rp 225.000'],
        ['13', 'Buck Converter LM2596', '12V→5V 3A', '1', 'Rp 35.000', 'Rp 35.000'],
        ['14', 'Buck Converter LM2596', '12V→3.3V 2A', '1', 'Rp 35.000', 'Rp 35.000'],
        ['', '', 'Subtotal B', '', '', 'Rp 625.000']
    ]
    add_table_with_data(doc, bom_b_headers, bom_b_data, 'Tabel 10.2: BOM Power Supply')
    
    doc.add_paragraph()
    doc.add_paragraph('C. Enclosure & Konektivitas', style='Heading 3')
    bom_c_headers = ['No', 'Komponen', 'Spesifikasi', 'Qty', 'Harga Satuan', 'Total']
    bom_c_data = [
        ['15', 'Junction Box IP68', '250x150x90mm ABS', '1', 'Rp 110.000', 'Rp 110.000'],
        ['16', 'Cable Gland M16', 'Waterproof', '6', 'Rp 8.000', 'Rp 48.000'],
        ['17', 'Cable Gland M20', 'Waterproof', '1', 'Rp 12.000', 'Rp 12.000'],
        ['18', 'Terminal Block', 'Screw 2-pin 5.08mm', '15', 'Rp 2.000', 'Rp 30.000'],
        ['19', 'Kabel Sensor Shielded', '4-core, 5 meter', '3', 'Rp 45.000', 'Rp 135.000'],
        ['20', 'Kabel Power', '18AWG red/black 3m', '1 set', 'Rp 25.000', 'Rp 25.000'],
        ['21', 'Pelampung PVC', '4 inch, 40cm + dop', '3', 'Rp 35.000', 'Rp 105.000'],
        ['22', 'Silicone Sealant', 'Waterproof', '1', 'Rp 25.000', 'Rp 25.000'],
        ['', '', 'Subtotal C', '', '', 'Rp 490.000']
    ]
    add_table_with_data(doc, bom_c_headers, bom_c_data, 'Tabel 10.3: BOM Enclosure')
    
    doc.add_paragraph()
    doc.add_paragraph('D. Komponen Pendukung', style='Heading 3')
    bom_d_headers = ['No', 'Komponen', 'Spesifikasi', 'Qty', 'Harga Satuan', 'Total']
    bom_d_data = [
        ['23', 'SIM Card', 'Telkomsel/Indosat', '1', 'Rp 25.000', 'Rp 25.000'],
        ['24', 'Antena GSM', 'External SMA', '1', 'Rp 15.000', 'Rp 15.000'],
        ['25', 'Kabel Jumper', 'Female-to-Female', '1 set', 'Rp 10.000', 'Rp 10.000'],
        ['26', 'Resistor Kit', '10kΩ, 1kΩ, 470Ω', '1 set', 'Rp 5.000', 'Rp 5.000'],
        ['27', 'Kapasitor Kit', '100µF, 10µF, 470µF', '10', 'Rp 1.000', 'Rp 10.000'],
        ['28', 'LED Indicator', '3mm RGB', '3', 'Rp 500', 'Rp 1.500'],
        ['29', 'Push Button', 'Tactile 6x6mm', '1', 'Rp 1.000', 'Rp 1.000'],
        ['30', 'PCB Perfboard', 'Double-sided 10x15cm', '2', 'Rp 12.000', 'Rp 24.000'],
        ['31', 'Kabel Tie', 'Nylon 20cm', '1 pack', 'Rp 15.000', 'Rp 15.000'],
        ['32', 'Label Waterproof', 'Identifikasi kabel', '1 set', 'Rp 20.000', 'Rp 20.000'],
        ['33', 'Fuse Holder + Fuse', '5A proteksi', '1 set', 'Rp 15.000', 'Rp 15.000'],
        ['34', 'TVS Diode', 'Proteksi spike', '3', 'Rp 2.000', 'Rp 6.000'],
        ['', '', 'Subtotal D', '', '', 'Rp 141.500']
    ]
    add_table_with_data(doc, bom_d_headers, bom_d_data, 'Tabel 10.4: BOM Komponen Pendukung')
    
    doc.add_paragraph()
    doc.add_heading('10.2 Biaya Jasa & Implementasi', level=2)
    jasa_headers = ['No', 'Item', 'Deskripsi', 'Total']
    jasa_data = [
        ['35', 'Jasa Engineering', 'Desain, assembly, programming', 'Rp 5.000.000'],
        ['36', 'Setup Cloud', 'ThingsBoard configuration', 'Rp 2.000.000'],
        ['37', 'ERP Integration', 'API development & testing', 'Rp 3.500.000'],
        ['38', 'Instalasi On-site', 'Pemasangan & kalibrasi', 'Rp 2.500.000'],
        ['39', 'Training', 'Operator training (2 hari)', 'Rp 1.500.000'],
        ['40', 'Dokumentasi', 'Manual, schematic, code', 'Rp 1.000.000'],
        ['', '', 'Subtotal Jasa', 'Rp 15.500.000']
    ]
    add_table_with_data(doc, jasa_headers, jasa_data, 'Tabel 10.5: Biaya Jasa')
    
    doc.add_paragraph()
    doc.add_heading('10.4 Rekapitulasi Biaya', level=2)
    rekap_headers = ['Kategori', 'Total']
    rekap_data = [
        ['A. Komponen Utama', 'Rp 16.011.000'],
        ['B. Power Supply', 'Rp 625.000'],
        ['C. Enclosure & Konektivitas', 'Rp 490.000'],
        ['D. Komponen Pendukung', 'Rp 141.500'],
        ['Subtotal Hardware', 'Rp 17.267.500'],
        ['Jasa & Implementasi', 'Rp 15.500.000'],
        ['TOTAL INVESTASI AWAL', 'Rp 32.767.500'],
        ['Biaya Operasional/Tahun', 'Rp 3.800.000']
    ]
    add_table_with_data(doc, rekap_headers, rekap_data, 'Tabel 10.6: Rekapitulasi Total')
    
    doc.add_paragraph()
    doc.add_heading('10.5 Opsi Paket', level=2)
    paket_headers = ['Paket', 'Deskripsi', 'Harga']
    paket_data = [
        ['Basic', '1 unit + instalasi + training', 'Rp 32.767.500'],
        ['Standard', '3 unit + instalasi + training + 1 tahun maintenance', 'Rp 95.000.000'],
        ['Enterprise', '5 unit + custom ERP integration + 2 tahun maintenance', 'Rp 175.000.000']
    ]
    add_table_with_data(doc, paket_headers, paket_data, 'Tabel 10.7: Opsi Paket')
    doc.add_paragraph('*Diskon volume tersedia untuk pembelian 5+ unit.*', style='Caption')
    
    # Section 13: Penutup
    doc.add_heading('13. PENUTUP', level=1)
    
    doc.add_heading('13.1 Kesimpulan', level=2)
    doc.add_paragraph(
        'Smart Water Quality Monitoring System (SWQMS) merupakan investasi strategis yang akan:'
    )
    conclusions = [
        'Meningkatkan efisiensi operasional melalui monitoring real-time 24/7',
        'Mengurangi risiko kerugian dengan sistem peringatan dini',
        'Mendukung keputusan berbasis data melalui integrasi ERP',
        'Memberikan competitive advantage melalui adopsi teknologi 4.0'
    ]
    for conclusion in conclusions:
        doc.add_paragraph(conclusion, style='List Number')
    
    doc.add_paragraph(
        'Dengan investasi awal Rp 32.767.500 dan biaya operasional Rp 3.800.000/tahun, '
        'sistem ini diproyeksikan memberikan ROI 6–9 bulan melalui penghematan operasional '
        'dan pencegahan kerugian.'
    )
    
    doc.add_heading('13.2 Call to Action', level=2)
    doc.add_paragraph('Kami mengundang Bapak/Ibu untuk:')
    actions = [
        'Diskusi teknis lebih lanjut dengan tim engineering kami',
        'Site survey untuk assessment lokasi instalasi',
        'Demo sistem (unit prototype tersedia)',
        'Negosiasi kontrak sesuai kebutuhan spesifik'
    ]
    for action in actions:
        doc.add_paragraph(action, style='List Number')
    
    doc.add_heading('13.3 Kontak', level=2)
    doc.add_paragraph('Tim Engineering IoT')
    doc.add_paragraph('Email: engineering@swqms.id')
    doc.add_paragraph('Phone: +62 812-XXXX-XXXX')
    doc.add_paragraph('Website: www.swqms.id')
    doc.add_paragraph('Alamat: [Alamat Perusahaan]')
    
    doc.add_paragraph()
    doc.add_paragraph('_' * 60)
    doc.add_paragraph()
    
    doc.add_paragraph(
        'Demikian proposal ini kami sampaikan. Besar harapan kami untuk dapat bekerja sama '
        'dan memberikan solusi terbaik bagi perusahaan Bapak/Ibu.',
        style='Intense Quote'
    )
    
    doc.add_paragraph()
    doc.add_paragraph('Hormat kami,')
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph('[Nama Lengkap]')
    doc.add_paragraph('Project Manager')
    doc.add_paragraph('Smart Water Quality Monitoring System')
    doc.add_paragraph('Tanggal: 10 September 2026')
    
    doc.add_paragraph()
    doc.add_paragraph(
        'Proposal ini berlaku hingga 30 hari sejak tanggal penerbitan. '
        'Harga dapat berubah sesuai kondisi pasar dan ketersediaan komponen.',
        style='Caption'
    )
    
    # Save document
    os.makedirs('output', exist_ok=True)
    output_path = 'output/PROPOSAL_SWQMS.docx'
    doc.save(output_path)
    print(f'✅ DOCX file created successfully: {output_path}')
    print(f'📄 File size: {os.path.getsize(output_path)} bytes')
    return output_path

if __name__ == '__main__':
    build_proposal()
