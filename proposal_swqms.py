#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to convert the SWQMS proposal markdown to DOCX format
"""

from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

def create_proposal_docx():
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Title
    title = doc.add_heading('PROPOSAL TEKNIS & KOMERSIAL', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle = doc.add_heading('SISTEM MONITORING KUALITAS AIR BERBASIS IoT', level=1)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle2 = doc.add_heading('Smart Water Quality Monitoring System (SWQMS)', level=2)
    subtitle2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle2.runs[0].italic = True
    
    doc.add_paragraph()  # Spacer
    
    # Submission info
    doc.add_paragraph('Diajukan Kepada:', style='Intense Quote')
    doc.add_paragraph('[Nama Perusahaan / Klien]', style='Intense Quote')
    doc.add_paragraph()
    
    doc.add_paragraph('Disusun Oleh:', style='Intense Quote')
    doc.add_paragraph('[Tim Engineering IoT]', style='Intense Quote')
    doc.add_paragraph()
    
    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    info_table.rows[0].cells[0].text = 'Tanggal:'
    info_table.rows[0].cells[1].text = '10 September 2026'
    info_table.rows[1].cells[0].text = 'Nomor Proposal:'
    info_table.rows[1].cells[1].text = 'SWQMS-2026-001'
    info_table.rows[2].cells[0].text = 'Versi:'
    info_table.rows[2].cells[1].text = '1.0'
    
    doc.add_page_break()
    
    # Table of Contents
    doc.add_heading('DAFTAR ISI', level=1)
    toc_items = [
        '1. Ringkasan Eksekutif',
        '2. Latar Belakang',
        '3. Tujuan Proyek',
        '4. Manfaat & Value Proposition',
        '5. Spesifikasi Teknis Sistem',
        '6. Fitur & Fungsionalitas',
        '7. Arsitektur Sistem',
        '8. Metodologi Implementasi',
        '9. Timeline & Roadmap',
        '10. Rencana Anggaran Biaya (RAB)',
        '11. Deliverables',
        '12. Garansi & Dukungan Pasca Implementasi',
        '13. Penutup'
    ]
    for item in toc_items:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()
    
    # Section 1: Ringkasan Eksekutif
    doc.add_heading('1. RINGKASAN EKSEKUTIF', level=1)
    
    p = doc.add_paragraph()
    p.add_run('Smart Water Quality Monitoring System (SWQMS)').bold = True
    p.add_run(' adalah solusi Internet of Things (IoT) terintegrasi untuk memantau kualitas air secara ')
    p.add_run('real-time').italic = True
    p.add_run(' pada kolam budidaya, reservoir, atau badan air lainnya. Sistem ini dirancang untuk beroperasi secara ')
    p.add_run('otonom 24/7').bold = True
    p.add_run(' dengan sumber daya mandiri (solar panel + baterai), mampu bekerja di kedalaman 1–3 meter, dan terintegrasi langsung dengan sistem ERP perusahaan.')
    
    doc.add_heading('Highlight Utama:', level=2)
    highlights = [
        '5 Parameter Kualitas Air: pH, EC/TDS/Salinitas, DO (Dissolved Oxygen), Suhu, dan monitoring kincir',
        'Otonom Energi: Solar panel 20W + baterai 12V 7.2Ah (backup 48+ jam)',
        'Dual Connectivity: WiFi (MQTT → ThingsBoard) + GSM (SMS Alert)',
        'Integrasi ERP: REST API untuk sinkronisasi data otomatis',
        'Smart Recommendation: AI-based saran perbaikan kualitas air & pakan',
        'Industrial Grade: Sensor IP68, casing waterproof, floating design'
    ]
    for highlight in highlights:
        doc.add_paragraph('✅ ' + highlight, style='List Bullet')
    
    p = doc.add_paragraph()
    p.add_run('Investasi Total: Rp 17.258.500 per unit').bold = True
    p.add_run(' dengan ROI (Return on Investment) diperkirakan ')
    p.add_run('6–9 bulan').bold = True
    p.add_run(' melalui efisiensi operasional dan pencegahan kerugian.')
    
    doc.add_page_break()
    
    # Section 2: Latar Belakang
    doc.add_heading('2. LATAR BELAKANG', level=1)
    
    doc.add_heading('2.1 Permasalahan yang Dihadapi', level=2)
    p = doc.add_paragraph()
    p.add_run('Dalam industri akuakultur dan pengelolaan sumber daya air, kualitas air merupakan ').bold = True
    p.add_run('faktor kritis').bold = True
    p.add_run(' yang menentukan:')
    
    factors = [
        'Kesehatan dan pertumbuhan organisme air (ikan, udang, dll)',
        'Efisiensi penggunaan pakan (FCR - Feed Conversion Ratio)',
        'Risiko kematian massal akibat penurunan kualitas air mendadak',
        'Kepatuhan terhadap regulasi lingkungan'
    ]
    for factor in factors:
        doc.add_paragraph(factor, style='List Bullet')
    
    doc.add_paragraph('Tantangan saat ini:', style='Intense Quote')
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
    doc.add_paragraph('Dengan perkembangan teknologi IoT, sensor industrial, dan cloud computing, kini memungkinkan untuk membangun sistem monitoring yang:')
    opportunities = [
        'Otonom (tidak bergantung listrik PLN)',
        'Real-time (data setiap 1–5 menit)',
        'Terintegrasi (cloud + ERP + mobile alert)',
        'Cerdas (AI-based recommendation)'
    ]
    for opp in opportunities:
        doc.add_paragraph(opp, style='List Bullet')
    
    doc.add_page_break()
    
    # Section 3: Tujuan Proyek
    doc.add_heading('3. TUJUAN PROYEK', level=1)
    
    doc.add_heading('3.1 Tujuan Umum', level=2)
    doc.add_paragraph('Membangun sistem monitoring kualitas air berbasis IoT yang handal, otonom, dan terintegrasi untuk mendukung operasional budidaya/perusahaan secara efisien dan berkelanjutan.')
    
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
    
    doc.add_page_break()
    
    # Section 4: Manfaat & Value Proposition
    doc.add_heading('4. MANFAAT & VALUE PROPOSITION', level=1)
    
    doc.add_heading('4.1 Manfaat Operasional', level=2)
    table = doc.add_table(rows=6, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Aspek'
    hdr_cells[1].text = 'Sebelum'
    hdr_cells[2].text = 'Sesudah'
    
    operational_data = [
        ('Monitoring', 'Manual, 1–2x/hari', 'Otomatis, setiap 1 menit'),
        ('Deteksi masalah', 'Reaktif (setelah kerugian)', 'Proaktif (peringatan dini)'),
        ('Data recording', 'Spreadsheet manual', 'Cloud database terstruktur'),
        ('Integrasi ERP', 'Tidak ada', 'REST API otomatis'),
        ('Respons time', 'Jam–hari', 'Menit (via SMS)')
    ]
    
    for i, (aspect, before, after) in enumerate(operational_data):
        table.rows[i+1].cells[0].text = aspect
        table.rows[i+1].cells[1].text = before
        table.rows[i+1].cells[2].text = after
    
    doc.add_heading('4.2 Manfaat Ekonomi', level=2)
    economic_benefits = [
        'Pengurangan kerugian akibat kematian ikan/udang: Rp 50–200 juta/tahun',
        'Efisiensi pakan melalui rekomendasi cerdas: 10–15% penghematan',
        'Pengurangan tenaga kerja monitoring manual: 1–2 orang',
        'ROI (Return on Investment): 6–9 bulan'
    ]
    for benefit in economic_benefits:
        doc.add_paragraph(benefit, style='List Bullet')
    
    doc.add_heading('4.3 Manfaat Strategis', level=2)
    strategic_benefits = [
        'Data-driven decision making untuk manajemen',
        'Kepatuhan regulasi lingkungan (data ter-audit)',
        'Competitive advantage melalui teknologi 4.0',
        'Sustainability melalui pengelolaan sumber daya optimal'
    ]
    for benefit in strategic_benefits:
        doc.add_paragraph(benefit, style='List Bullet')
    
    doc.add_page_break()
    
    # Section 5: Spesifikasi Teknis Sistem
    doc.add_heading('5. SPESIFIKASI TEKNIS SISTEM', level=1)
    
    doc.add_heading('5.1 Spesifikasi Sensor', level=2)
    sensor_table = doc.add_table(rows=6, cols=7)
    sensor_table.style = 'Table Grid'
    hdr_cells = sensor_table.rows[0].cells
    headers = ['No', 'Parameter', 'Model', 'Range', 'Akurasi', 'Output', 'Proteksi']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    sensor_data = [
        ('1', 'pH', 'Industrial RS485', '0–14 pH', '±0.1 pH', 'Modbus RTU', 'IP68'),
        ('2', 'EC/TDS/Salinitas', 'Industrial RS485', '0–20000 µS/cm', '±1%', 'Modbus RTU', 'IP68'),
        ('3', 'Dissolved Oxygen', 'Optical RS485', '0–20 mg/L', '±0.1 mg/L', 'Modbus RTU', 'IP68'),
        ('4', 'Suhu', 'PT100 RS485', '-10–85°C', '±0.5°C', 'Modbus RTU', 'IP68'),
        ('5', 'Arus Kincir', 'SCT-013-030', '0–30A AC', '±1%', 'Analog 0–1V', 'Non-invasive')
    ]
    
    for i, data in enumerate(sensor_data):
        for j, value in enumerate(data):
            sensor_table.rows[i+1].cells[j].text = value
    
    doc.add_heading('5.2 Spesifikasi Mikrokontroler & Komunikasi', level=2)
    mcu_table = doc.add_table(rows=7, cols=2)
    mcu_table.style = 'Table Grid'
    mcu_data = [
        ('MCU', 'ESP32 DevKit V1 (Dual-core 240MHz, WiFi 802.11 b/g/n, Bluetooth 4.2)'),
        ('Memory', '520KB SRAM, 4MB Flash'),
        ('Protokol', 'MQTT (ThingsBoard), HTTP/REST (ERP), AT Command (SMS)'),
        ('GSM Module', 'SIM800L Quad-band (850/900/1800/1900 MHz)'),
        ('RS485', 'MAX3485, 3.3V compatible, auto-direction'),
        ('Jarak Komunikasi', 'RS485: hingga 1200m, WiFi: 100m, GSM: nationwide')
    ]
    for i, (component, spec) in enumerate(mcu_data):
        mcu_table.rows[i].cells[0].text = component
        mcu_table.rows[i].cells[1].text = spec
    
    doc.add_heading('5.3 Spesifikasi Power Supply', level=2)
    power_table = doc.add_table(rows=7, cols=2)
    power_table.style = 'Table Grid'
    power_data = [
        ('Solar Panel', 'Monocrystalline 20W, 12V, waterproof'),
        ('Charge Controller', 'PWM 10A, 12V/24V auto-detect, LCD display'),
        ('Battery', 'VRLA/SLA 12V 7.2Ah, maintenance-free'),
        ('Backup Time', '48+ jam tanpa sinar matahari'),
        ('Regulator', '2x LM2596 Buck Converter (5V 3A & 3.3V 2A)'),
        ('Proteksi', 'Fuse 5A, TVS Diode, overcharge/overdischarge protection')
    ]
    for i, (component, spec) in enumerate(power_data):
        power_table.rows[i].cells[0].text = component
        power_table.rows[i].cells[1].text = spec
    
    doc.add_heading('5.4 Spesifikasi Enclosure & Mekanikal', level=2)
    enclosure_table = doc.add_table(rows=9, cols=2)
    enclosure_table.style = 'Table Grid'
    enclosure_data = [
        ('Casing', 'Junction Box IP68, ABS plastic, 250x150x90mm'),
        ('Cable Gland', '6x M16 + 1x M20, waterproof'),
        ('Pelampung', '3x PVC pipe 4 inch, panjang 40cm'),
        ('Kabel Sensor', 'Shielded 4-core, 5 meter, UV resistant'),
        ('Operating Temp', '-20°C to +60°C'),
        ('Humidity', '0–100% RH (condensing)'),
        ('Dimensi Total', '±60cm x 40cm x 30cm (termasuk pelampung)'),
        ('Berat Total', '±4.5 kg (termasuk baterai)')
    ]
    for i, (component, spec) in enumerate(enclosure_data):
        enclosure_table.rows[i].cells[0].text = component
        enclosure_table.rows[i].cells[1].text = spec
    
    doc.add_page_break()
    
    # Section 6: Fitur & Fungsionalitas
    doc.add_heading('6. FITUR & FUNGSIONALITAS', level=1)
    
    doc.add_heading('6.1 Fitur Utama', level=2)
    
    doc.add_heading('1. Monitoring Multi-Parameter Real-Time', level=3)
    doc.add_paragraph('Pembacaan 5 parameter kualitas air setiap 1–5 menit')
    doc.add_paragraph('Data dikirim ke cloud via MQTT protocol (hemat bandwidth)')
    doc.add_paragraph('Dashboard visualisasi real-time di ThingsBoard')
    
    doc.add_heading('2. Sistem Daya Otonom', level=3)
    doc.add_paragraph('Solar panel 20W untuk charging siang hari')
    doc.add_paragraph('Baterai 12V 7.2Ah untuk backup 48+ jam')
    doc.add_paragraph('Manajemen daya cerdas: deep sleep mode antar pembacaan')
    doc.add_paragraph('Monitoring tegangan baterai real-time')
    
    doc.add_heading('3. Dual Connectivity', level=3)
    doc.add_paragraph('WiFi: Untuk transmisi data rutin ke ThingsBoard')
    doc.add_paragraph('GSM (SIM800L): Fallback jika WiFi down + SMS alert kritis')
    doc.add_paragraph('Auto-switching antar koneksi')
    
    doc.add_heading('4. Sistem Peringatan Dini (Early Warning System)', level=3)
    doc.add_paragraph('SMS Alert ke nomor yang terdaftar saat:')
    alerts = [
        'Parameter melebihi threshold (pH, DO, suhu, dll)',
        'Sistem down / sensor error',
        'Kincir mati terdeteksi',
        'Baterai low (<20%)'
    ]
    for alert in alerts:
        doc.add_paragraph(alert, style='List Bullet')
    doc.add_paragraph('Email/WhatsApp notification (opsional via API)')
    
    doc.add_heading('5. Deteksi Kegagalan Peralatan', level=3)
    doc.add_paragraph('SCT-013 current sensor untuk monitoring kincir/aerator')
    doc.add_paragraph('Deteksi kincir mati dalam 30 detik')
    doc.add_paragraph('Alert otomatis ke operator')
    
    doc.add_heading('6. Cloud Storage & ERP Integration', level=3)
    doc.add_paragraph('ThingsBoard Cloud: Penyimpanan data historis unlimited')
    doc.add_paragraph('REST API: Integrasi dengan sistem ERP perusahaan')
    doc.add_paragraph('Data export: CSV, JSON, XML untuk analisis')
    doc.add_paragraph('Multi-device support: Scalable untuk banyak unit')
    
    doc.add_heading('7. Smart Recommendation Engine', level=3)
    doc.add_paragraph('Rule-based system untuk saran perbaikan kualitas air')
    doc.add_paragraph('Rekomendasi pakan berdasarkan kondisi air')
    doc.add_paragraph('Prediksi tren menggunakan data historis')
    doc.add_paragraph('Notifikasi proaktif sebelum masalah terjadi')
    
    doc.add_heading('8. Dashboard & Analytics', level=3)
    doc.add_paragraph('Real-time dashboard dengan widget interaktif')
    doc.add_paragraph('Historical charts (harian, mingguan, bulanan)')
    doc.add_paragraph('Export report otomatis (PDF/Excel)')
    doc.add_paragraph('Multi-user access dengan role-based permission')
    
    doc.add_heading('6.2 Fitur Tambahan (Opsional)', level=2)
    optional_table = doc.add_table(rows=6, cols=3)
    optional_table.style = 'Table Grid'
    hdr_cells = optional_table.rows[0].cells
    hdr_cells[0].text = 'Fitur'
    hdr_cells[1].text = 'Deskripsi'
    hdr_cells[2].text = 'Biaya Tambahan'
    
    optional_data = [
        ('Camera Module', 'Monitoring visual kolam', 'Rp 350.000'),
        ('Rain Sensor', 'Deteksi curah hujan', 'Rp 85.000'),
        ('Water Level Sensor', 'Monitoring ketinggian air', 'Rp 120.000'),
        ('Local Data Logger', 'SD card backup', 'Rp 95.000'),
        ('LCD Display', 'Display lokal di unit', 'Rp 150.000')
    ]
    
    for i, (feature, desc, cost) in enumerate(optional_data):
        optional_table.rows[i+1].cells[0].text = feature
        optional_table.rows[i+1].cells[1].text = desc
        optional_table.rows[i+1].cells[2].text = cost
    
    doc.add_page_break()
    
    # Section 7: Arsitektur Sistem
    doc.add_heading('7. ARSITEKTUR SISTEM', level=1)
    
    doc.add_heading('7.1 Diagram Arsitektur High-Level', level=2)
    doc.add_paragraph('Diagram arsitektur sistem terdiri dari 4 layer:', style='Intense Quote')
    doc.add_paragraph('Layer 1: Edge Device (Field) - ESP32 dengan sensor dan modul komunikasi', style='List Bullet')
    doc.add_paragraph('Layer 2: Cloud - ThingsBoard untuk telemetry dan dashboard', style='List Bullet')
    doc.add_paragraph('Layer 3: Notification - SMS Gateway via SIM800L', style='List Bullet')
    doc.add_paragraph('Layer 4: Integration - ERP System via REST API', style='List Bullet')
    
    doc.add_heading('7.2 Alur Data (Data Flow)', level=2)
    doc.add_paragraph('Sensor → RS485 → ESP32 → WiFi → MQTT → ThingsBoard', style='Quote')
    doc.add_paragraph('                              ↓', style='Quote')
    doc.add_paragraph('                         GSM → SMS Alert', style='Quote')
    doc.add_paragraph('                              ↓', style='Quote')
    doc.add_paragraph('                    REST API → ERP System', style='Quote')
    
    doc.add_heading('7.3 Protokol Komunikasi', level=2)
    protocol_table = doc.add_table(rows=5, cols=4)
    protocol_table.style = 'Table Grid'
    hdr_cells = protocol_table.rows[0].cells
    headers = ['Layer', 'Protokol', 'Port', 'Enkripsi']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    protocol_data = [
        ('Sensor → ESP32', 'Modbus RTU (RS485)', '-', '-'),
        ('ESP32 → Cloud', 'MQTT over TLS', '8883', 'TLS 1.2'),
        ('ESP32 → SMS', 'AT Commands (GSM)', '-', '-'),
        ('Cloud → ERP', 'HTTPS REST API', '443', 'TLS 1.2')
    ]
    
    for i, data in enumerate(protocol_data):
        for j, value in enumerate(data):
            protocol_table.rows[i+1].cells[j].text = value
    
    doc.add_page_break()
    
    # Section 8: Metodologi Implementasi
    doc.add_heading('8. METODOLOGI IMPLEMENTASI', level=1)
    
    doc.add_heading('8.1 Pendekatan Proyek', level=2)
    doc.add_paragraph('Menggunakan metodologi Agile-Waterfall Hybrid:')
    doc.add_paragraph('Waterfall untuk fase desain & procurement (terstruktur)', style='List Bullet')
    doc.add_paragraph('Agile untuk fase development & testing (iteratif)', style='List Bullet')
    
    doc.add_heading('8.2 Tahapan Implementasi', level=2)
    
    phases = [
        ('Fase 1: Discovery & Design (Minggu 1–2)', [
            'Survey lokasi instalasi',
            'Finalisasi spesifikasi teknis',
            'Desain arsitektur sistem',
            'Approval desain oleh klien'
        ]),
        ('Fase 2: Procurement (Minggu 3–4)', [
            'Pembelian komponen (BOM)',
            'Quality check komponen',
            'Preparasi workshop'
        ]),
        ('Fase 3: Development & Assembly (Minggu 5–7)', [
            'Assembly hardware (PCB, wiring)',
            'Development firmware ESP32',
            'Setup ThingsBoard cloud',
            'Konfigurasi Rule Engine',
            'Integrasi ERP API'
        ]),
        ('Fase 4: Testing & QA (Minggu 8–9)', [
            'Unit testing per modul',
            'Integration testing',
            'Waterproof testing (IP68)',
            'Field testing (kolam nyata)',
            'Stress testing (48 jam non-stop)'
        ]),
        ('Fase 5: Deployment (Minggu 10)', [
            'Instalasi di lokasi',
            'Kalibrasi sensor',
            'Training operator',
            'Handover dokumentasi'
        ]),
        ('Fase 6: Monitoring & Support (Minggu 11–12)', [
            'Monitoring performa',
            'Bug fixing',
            'Optimasi sistem',
            'Final report'
        ])
    ]
    
    for phase_name, tasks in phases:
        doc.add_heading(phase_name, level=3)
        for task in tasks:
            doc.add_paragraph(task, style='List Bullet')
    
    doc.add_page_break()
    
    # Section 9: Timeline & Roadmap
    doc.add_heading('9. TIMELINE & ROADMAP', level=1)
    
    doc.add_heading('9.1 Gantt Chart Proyek', level=2)
    doc.add_paragraph('Durasi total proyek: 12 minggu', style='Intense Quote')
    
    milestones = [
        'M1: Design Approved (Minggu 2)',
        'M2: Components Ready (Minggu 4)',
        'M3: System Integrated (Minggu 7)',
        'M4: Testing Passed (Minggu 9)',
        'M5: Deployed (Minggu 10)',
        'M6: Project Closed (Minggu 12)'
    ]
    for milestone in milestones:
        doc.add_paragraph(milestone, style='List Bullet')
    
    doc.add_heading('9.2 Roadmap Pengembangan (Post-Deployment)', level=2)
    roadmap_table = doc.add_table(rows=5, cols=2)
    roadmap_table.style = 'Table Grid'
    hdr_cells = roadmap_table.rows[0].cells
    hdr_cells[0].text = 'Kuartal'
    hdr_cells[1].text = 'Pengembangan'
    
    roadmap_data = [
        ('Q1 2027', 'Multi-unit deployment (5–10 unit)'),
        ('Q2 2027', 'Machine Learning model untuk prediksi'),
        ('Q3 2027', 'Mobile app native (iOS/Android)'),
        ('Q4 2027', 'Integrasi dengan sistem feeding otomatis')
    ]
    
    for i, (quarter, dev) in enumerate(roadmap_data):
        roadmap_table.rows[i+1].cells[0].text = quarter
        roadmap_table.rows[i+1].cells[1].text = dev
    
    doc.add_page_break()
    
    # Section 10: RAB
    doc.add_heading('10. RENCANA ANGGARAN BIAYA (RAB)', level=1)
    
    doc.add_heading('10.1 Bill of Materials (BOM)', level=2)
    
    doc.add_heading('A. Komponen Utama (Sensor & Mikrokontroler)', level=3)
    bom_a_table = doc.add_table(rows=10, cols=6)
    bom_a_table.style = 'Table Grid'
    hdr_cells = bom_a_table.rows[0].cells
    headers = ['No', 'Komponen', 'Spesifikasi', 'Qty', 'Harga Satuan', 'Total']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    bom_a_data = [
        ('1', 'ESP32 DevKit V1', 'ESP32-WROOM-32, WiFi+BT', '1', 'Rp 75.000', 'Rp 75.000'),
        ('2', 'Sensor pH Industrial', 'RS485 Modbus, IP68, kabel 5m', '1', 'Rp 3.150.000', 'Rp 3.150.000'),
        ('3', 'Sensor EC/TDS/Salinity', 'RS485 Modbus, IP68, kabel 5m', '1', 'Rp 3.650.000', 'Rp 3.650.000'),
        ('4', 'Sensor DO Industrial', 'Optical RS485, IP68, kabel 5m', '1', 'Rp 8.500.000', 'Rp 8.500.000'),
        ('5', 'Sensor Suhu Industrial', 'PT100 RS485, IP68, kabel 5m', '1', 'Rp 450.000', 'Rp 450.000'),
        ('6', 'Modul RS485 to TTL', 'MAX3485, 3.3V compatible', '1', 'Rp 25.000', 'Rp 25.000'),
        ('7', 'SIM800L GSM Module', 'Quad-band, SMS+GPRS', '1', 'Rp 65.000', 'Rp 65.000'),
        ('8', 'SCT-013-030', 'Current Transformer 30A', '1', 'Rp 45.000', 'Rp 45.000'),
        ('9', 'Burden Resistor 62Ω', '1% tolerance', '1', 'Rp 1.000', 'Rp 1.000'),
        ('', '', '', '', 'Subtotal A:', 'Rp 16.011.000')
    ]
    
    for i, data in enumerate(bom_a_data):
        for j, value in enumerate(data):
            bom_a_table.rows[i].cells[j].text = value
    
    doc.add_heading('B. Power Supply', level=3)
    bom_b_table = doc.add_table(rows=7, cols=6)
    bom_b_table.style = 'Table Grid'
    hdr_cells = bom_b_table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    bom_b_data = [
        ('10', 'Solar Panel', 'Monocrystalline 20W 12V', '1', 'Rp 220.000', 'Rp 220.000'),
        ('11', 'Solar Charge Controller', 'PWM 10A, LCD display', '1', 'Rp 110.000', 'Rp 110.000'),
        ('12', 'Baterai Kering', 'VRLA/SLA 12V 7.2Ah', '1', 'Rp 225.000', 'Rp 225.000'),
        ('13', 'Buck Converter LM2596', '12V→5V 3A', '1', 'Rp 35.000', 'Rp 35.000'),
        ('14', 'Buck Converter LM2596', '12V→3.3V 2A', '1', 'Rp 35.000', 'Rp 35.000'),
        ('', '', '', '', 'Subtotal B:', 'Rp 625.000')
    ]
    
    for i, data in enumerate(bom_b_data):
        for j, value in enumerate(data):
            bom_b_table.rows[i].cells[j].text = value
    
    doc.add_heading('C. Enclosure & Konektivitas', level=3)
    bom_c_table = doc.add_table(rows=10, cols=6)
    bom_c_table.style = 'Table Grid'
    hdr_cells = bom_c_table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    bom_c_data = [
        ('15', 'Junction Box IP68', '250x150x90mm ABS', '1', 'Rp 110.000', 'Rp 110.000'),
        ('16', 'Cable Gland M16', 'Waterproof', '6', 'Rp 8.000', 'Rp 48.000'),
        ('17', 'Cable Gland M20', 'Waterproof', '1', 'Rp 12.000', 'Rp 12.000'),
        ('18', 'Terminal Block', 'Screw 2-pin 5.08mm', '15', 'Rp 2.000', 'Rp 30.000'),
        ('19', 'Kabel Sensor Shielded', '4-core, 5 meter', '3', 'Rp 45.000', 'Rp 135.000'),
        ('20', 'Kabel Power', '18AWG red/black 3m', '1 set', 'Rp 25.000', 'Rp 25.000'),
        ('21', 'Pelampung PVC', '4 inch, 40cm + dop', '3', 'Rp 35.000', 'Rp 105.000'),
        ('22', 'Silicone Sealant', 'Waterproof', '1', 'Rp 25.000', 'Rp 25.000'),
        ('', '', '', '', 'Subtotal C:', 'Rp 490.000')
    ]
    
    for i, data in enumerate(bom_c_data):
        for j, value in enumerate(data):
            bom_c_table.rows[i].cells[j].text = value
    
    doc.add_heading('D. Komponen Pendukung', level=3)
    bom_d_table = doc.add_table(rows=14, cols=6)
    bom_d_table.style = 'Table Grid'
    hdr_cells = bom_d_table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
    
    bom_d_data = [
        ('23', 'SIM Card', 'Telkomsel/Indosat', '1', 'Rp 25.000', 'Rp 25.000'),
        ('24', 'Antena GSM', 'External SMA', '1', 'Rp 15.000', 'Rp 15.000'),
        ('25', 'Kabel Jumper', 'Female-to-Female', '1 set', 'Rp 10.000', 'Rp 10.000'),
        ('26', 'Resistor Kit', '10kΩ, 1kΩ, 470Ω', '1 set', 'Rp 5.000', 'Rp 5.000'),
        ('27', 'Kapasitor Kit', '100µF, 10µF, 470µF', '10', 'Rp 1.000', 'Rp 10.000'),
        ('28', 'LED Indicator', '3mm RGB', '3', 'Rp 500', 'Rp 1.500'),
        ('29', 'Push Button', 'Tactile 6x6mm', '1', 'Rp 1.000', 'Rp 1.000'),
        ('30', 'PCB Perfboard', 'Double-sided 10x15cm', '2', 'Rp 12.000', 'Rp 24.000'),
        ('31', 'Kabel Tie', 'Nylon 20cm', '1 pack', 'Rp 15.000', 'Rp 15.000'),
        ('32', 'Label Waterproof', 'Identifikasi kabel', '1 set', 'Rp 20.000', 'Rp 20.000'),
        ('33', 'Fuse Holder + Fuse', '5A proteksi', '1 set', 'Rp 15.000', 'Rp 15.000'),
        ('34', 'TVS Diode', 'Proteksi spike', '3', 'Rp 2.000', 'Rp 6.000'),
        ('', '', '', '', 'Subtotal D:', 'Rp 141.500')
    ]
    
    for i, data in enumerate(bom_d_data):
        for j, value in enumerate(data):
            bom_d_table.rows[i].cells[j].text = value
    
    doc.add_heading('10.2 Biaya Jasa & Implementasi', level=2)
    jasa_table = doc.add_table(rows=7, cols=4)
    jasa_table.style = 'Table Grid'
    hdr_cells = jasa_table.rows[0].cells
    headers_jasa = ['No', 'Item', 'Deskripsi', 'Total']
    for i, header in enumerate(headers_jasa):
        hdr_cells[i].text = header
    
    jasa_data = [
        ('35', 'Jasa Engineering', 'Desain, assembly, programming', 'Rp 5.000.000'),
        ('36', 'Setup Cloud', 'ThingsBoard configuration', 'Rp 2.000.000'),
        ('37', 'ERP Integration', 'API development & testing', 'Rp 3.500.000'),
        ('38', 'Instalasi On-site', 'Pemasangan & kalibrasi', 'Rp 2.500.000'),
        ('39', 'Training', 'Operator training (2 hari)', 'Rp 1.500.000'),
        ('40', 'Dokumentasi', 'Manual, schematic, code', 'Rp 1.000.000'),
        ('', '', 'Subtotal Jasa:', 'Rp 15.500.000')
    ]
    
    for i, data in enumerate(jasa_data):
        for j, value in enumerate(data):
            jasa_table.rows[i].cells[j].text = value
    
    doc.add_heading('10.3 Biaya Operasional (Tahunan)', level=2)
    ops_table = doc.add_table(rows=5, cols=4)
    ops_table.style = 'Table Grid'
    hdr_cells = ops_table.rows[0].cells
    headers_ops = ['No', 'Item', 'Deskripsi', 'Biaya/Tahun']
    for i, header in enumerate(headers_ops):
        hdr_cells[i].text = header
    
    ops_data = [
        ('41', 'Kuota Data GSM', '1GB/bulan (SMS + backup)', 'Rp 600.000'),
        ('42', 'ThingsBoard Cloud', 'Professional plan', 'Rp 1.200.000'),
        ('43', 'Maintenance', 'Kalibrasi sensor 2x/tahun', 'Rp 2.000.000'),
        ('', '', 'Subtotal Operasional:', 'Rp 3.800.000')
    ]
    
    for i, data in enumerate(ops_data):
        for j, value in enumerate(data):
            ops_table.rows[i].cells[j].text = value
    
    doc.add_heading('10.4 Rekapitulasi Biaya', level=2)
    rekap_table = doc.add_table(rows=8, cols=2)
    rekap_table.style = 'Table Grid'
    rekap_data = [
        ('A. Komponen Utama', 'Rp 16.011.000'),
        ('B. Power Supply', 'Rp 625.000'),
        ('C. Enclosure & Konektivitas', 'Rp 490.000'),
        ('D. Komponen Pendukung', 'Rp 141.500'),
        ('Subtotal Hardware', 'Rp 17.267.500'),
        ('Jasa & Implementasi', 'Rp 15.500.000'),
        ('TOTAL INVESTASI AWAL', 'Rp 32.767.500'),
        ('Biaya Operasional/Tahun', 'Rp 3.800.000')
    ]
    
    for i, (category, total) in enumerate(rekap_data):
        rekap_table.rows[i].cells[0].text = category
        rekap_table.rows[i].cells[1].text = total
    
    doc.add_heading('10.5 Opsi Paket', level=2)
    paket_table = doc.add_table(rows=4, cols=3)
    paket_table.style = 'Table Grid'
    hdr_cells = paket_table.rows[0].cells
    headers_paket = ['Paket', 'Deskripsi', 'Harga']
    for i, header in enumerate(headers_paket):
        hdr_cells[i].text = header
    
    paket_data = [
        ('Basic', '1 unit + instalasi + training', 'Rp 32.767.500'),
        ('Standard', '3 unit + instalasi + training + 1 tahun maintenance', 'Rp 95.000.000'),
        ('Enterprise', '5 unit + custom ERP integration + 2 tahun maintenance', 'Rp 175.000.000')
    ]
    
    for i, (paket, desc, price) in enumerate(paket_data):
        paket_table.rows[i+1].cells[0].text = paket
        paket_table.rows[i+1].cells[1].text = desc
        paket_table.rows[i+1].cells[2].text = price
    
    doc.add_paragraph('*Diskon volume tersedia untuk pembelian 5+ unit.', style='Intense Quote')
    
    doc.add_page_break()
    
    # Section 11: Deliverables
    doc.add_heading('11. DELIVERABLES', level=1)
    
    doc.add_heading('11.1 Hardware', level=2)
    hardware_deliverables = [
        '1 unit Smart Water Quality Monitoring System (lengkap)',
        '5 sensor probe industrial (pH, EC/TDS, DO, Suhu, Current)',
        'Solar panel 20W dengan mounting bracket',
        'Kabel sensor 5 meter (shielded)',
        'Spare parts kit (fuse, resistor, kabel tie)'
    ]
    for item in hardware_deliverables:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('11.2 Software & Cloud', level=2)
    software_deliverables = [
        'Firmware ESP32 (source code + binary)',
        'ThingsBoard dashboard (customized)',
        'Rule Engine configuration',
        'REST API documentation',
        'ERP integration module'
    ]
    for item in software_deliverables:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('11.3 Dokumentasi', level=2)
    doc_deliverables = [
        'User Manual (Bahasa Indonesia)',
        'Technical Schematic (PDF + source file)',
        'BOM (Bill of Materials) lengkap',
        'Installation Guide',
        'Maintenance Manual',
        'API Documentation'
    ]
    for item in doc_deliverables:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('11.4 Training', level=2)
    training_deliverables = [
        'Training operator (2 hari)',
        'Training admin dashboard (1 hari)',
        'Video tutorial (rekaman)',
        'Sertifikat training'
    ]
    for item in training_deliverables:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()
    
    # Section 12: Garansi & Dukungan
    doc.add_heading('12. GARANSI & DUKUNGAN PASCA IMPLEMENTASI', level=1)
    
    doc.add_heading('12.1 Garansi Hardware', level=2)
    warranty_table = doc.add_table(rows=7, cols=3)
    warranty_table.style = 'Table Grid'
    hdr_cells = warranty_table.rows[0].cells
    headers_warranty = ['Komponen', 'Masa Garansi', 'Coverage']
    for i, header in enumerate(headers_warranty):
        hdr_cells[i].text = header
    
    warranty_data = [
        ('ESP32 DevKit', '12 bulan', 'Defect pabrik'),
        ('Sensor Industrial', '12 bulan', 'Defect pabrik'),
        ('Solar Panel', '24 bulan', 'Defect + degradasi >20%'),
        ('Battery', '6 bulan', 'Defect pabrik'),
        ('Enclosure IP68', '12 bulan', 'Kebocoran'),
        ('Other Components', '6 bulan', 'Defect pabrik')
    ]
    
    for i, data in enumerate(warranty_data):
        for j, value in enumerate(data):
            warranty_table.rows[i+1].cells[j].text = value
    
    doc.add_heading('12.2 Dukungan Teknis', level=2)
    support_table = doc.add_table(rows=4, cols=4)
    support_table.style = 'Table Grid'
    hdr_cells = support_table.rows[0].cells
    headers_support = ['Paket Support', 'Response Time', 'Channel', 'Biaya']
    for i, header in enumerate(headers_support):
        hdr_cells[i].text = header
    
    support_data = [
        ('Basic', '48 jam', 'Email', 'Included 3 bulan'),
        ('Standard', '24 jam', 'Email + Phone', 'Rp 500.000/bulan'),
        ('Premium', '4 jam', 'Email + Phone + On-site', 'Rp 1.500.000/bulan')
    ]
    
    for i, data in enumerate(support_data):
        for j, value in enumerate(data):
            support_table.rows[i+1].cells[j].text = value
    
    doc.add_heading('12.3 Maintenance Preventif', level=2)
    doc.add_paragraph('Paket Maintenance Tahunan (Rp 2.000.000/tahun):', style='Intense Quote')
    maintenance_items = [
        'Kalibrasi sensor 2x per tahun',
        'Pembersihan probe sensor',
        'Cek waterproof casing',
        'Update firmware',
        'Backup data',
        'Report kesehatan sistem'
    ]
    for item in maintenance_items:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('12.4 SLA (Service Level Agreement)', level=2)
    sla_table = doc.add_table(rows=6, cols=2)
    sla_table.style = 'Table Grid'
    hdr_cells = sla_table.rows[0].cells
    headers_sla = ['Metrik', 'Target']
    for i, header in enumerate(headers_sla):
        hdr_cells[i].text = header
    
    sla_data = [
        ('System Uptime', '≥ 99%'),
        ('Data Accuracy', '≥ 98%'),
        ('Alert Delivery', '< 5 menit'),
        ('Support Response', '< 24 jam (Standard)'),
        ('On-site Visit', '< 3 hari kerja')
    ]
    
    for i, data in enumerate(sla_data):
        sla_table.rows[i+1].cells[0].text = data[0]
        sla_table.rows[i+1].cells[1].text = data[1]
    
    doc.add_page_break()
    
    # Section 13: Penutup
    doc.add_heading('13. PENUTUP', level=1)
    
    doc.add_heading('13.1 Kesimpulan', level=2)
    doc.add_paragraph('Smart Water Quality Monitoring System (SWQMS) merupakan investasi strategis yang akan:')
    conclusions = [
        'Meningkatkan efisiensi operasional melalui monitoring real-time 24/7',
        'Mengurangi risiko kerugian dengan sistem peringatan dini',
        'Mendukung keputusan berbasis data melalui integrasi ERP',
        'Memberikan competitive advantage melalui adopsi teknologi 4.0'
    ]
    for conclusion in conclusions:
        doc.add_paragraph(conclusion, style='List Number')
    
    p = doc.add_paragraph()
    p.add_run('Dengan ').bold = False
    p.add_run('investasi awal Rp 32.767.500').bold = True
    p.add_run(' dan ').bold = False
    p.add_run('biaya operasional Rp 3.800.000/tahun').bold = True
    p.add_run(', sistem ini diproyeksikan memberikan ').bold = False
    p.add_run('ROI 6–9 bulan').bold = True
    p.add_run(' melalui penghematan operasional dan pencegahan kerugian.')
    
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
    doc.add_paragraph('Tim Engineering IoT', style='Intense Quote')
    contact_info = [
        'Email: engineering@swqms.id',
        'Phone: +62 812-XXXX-XXXX',
        'Website: www.swqms.id',
        'Alamat: [Alamat Perusahaan]'
    ]
    for info in contact_info:
        doc.add_paragraph(info, style='List Bullet')
    
    doc.add_page_break()
    
    # Lampiran
    doc.add_heading('LAMPIRAN', level=1)
    
    doc.add_heading('Lampiran A: Datasheet Sensor', level=2)
    lampiran_a = [
        'Datasheet Sensor pH Industrial RS485',
        'Datasheet Sensor EC/TDS Industrial',
        'Datasheet Sensor DO Optical',
        'Datasheet Sensor Suhu PT100'
    ]
    for item in lampiran_a:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('Lampiran B: Schematic Diagram', level=2)
    lampiran_b = [
        'Diagram blok sistem lengkap',
        'Wiring diagram RS485 bus',
        'Power supply schematic',
        'ESP32 pin connection'
    ]
    for item in lampiran_b:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('Lampiran C: Sample Dashboard', level=2)
    lampiran_c = [
        'Screenshot ThingsBoard dashboard',
        'Sample report PDF',
        'Mobile view preview'
    ]
    for item in lampiran_c:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('Lampiran D: Referensi Proyek', level=2)
    lampiran_d = [
        'Case study implementasi serupa',
        'Testimoni klien',
        'Sertifikasi & standar'
    ]
    for item in lampiran_d:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading('Lampiran E: Terms & Conditions', level=2)
    lampiran_e = [
        'Syarat pembayaran',
        'Syarat garansi',
        'Ketentuan maintenance',
        'Kebijakan privasi data'
    ]
    for item in lampiran_e:
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_page_break()
    
    # Closing
    closing_text = """Demikian proposal ini kami sampaikan. Besar harapan kami untuk dapat bekerja sama dan memberikan solusi terbaik bagi perusahaan Bapak/Ibu.

Hormat kami,


[Tanda Tangan]


[Nama Lengkap]
Project Manager
Smart Water Quality Monitoring System
Tanggal: 10 September 2026


Proposal ini berlaku hingga 30 hari sejak tanggal penerbitan. Harga dapat berubah sesuai kondisi pasar dan ketersediaan komponen."""
    
    doc.add_paragraph(closing_text)
    
    # Save document
    doc.save('/workspace/PROPOSAL_SWQMS.docx')
    print("✅ Document successfully created: PROPOSAL_SWQMS.docx")

if __name__ == '__main__':
    create_proposal_docx()
