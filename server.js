const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Konfigurasi Multer untuk penampungan sementara file Excel
const upload = multer({ dest: 'uploads/' });

// Lokasi penyimpanan data arsip JSON
const DATA_FILE = path.join(__dirname, 'data', 'rekap_data.json');

// Pastikan folder data dan uploads ada
if (!fs.existsSync(path.join(__dirname, 'data'))) {
    fs.mkdirSync(path.join(__dirname, 'data'));
}
if (!fs.existsSync(path.join(__dirname, 'uploads'))) {
    fs.mkdirSync(path.join(__dirname, 'uploads'));
}
if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify([]), 'utf-8');
}

// Helper untuk membaca file JSON
function getSavedData() {
    try {
        const raw = fs.readFileSync(DATA_FILE, 'utf-8');
        return JSON.parse(raw);
    } catch (e) {
        return [];
    }
}

// Helper untuk menyimpan file JSON
function saveAllData(data) {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// 1. Endpoint Ambil Seluruh Data Arsip
app.get('/api/data', (req, res) => {
    const data = getSavedData();
    res.json({ success: true, data });
});

// 2. Endpoint Upload File Excel (Logika Overwrite + Auto Clean Temp File)
app.post('/upload', upload.single('excelFile'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'File belum dipilih!' });
        }

        const { bulan, gelombang, golongan, jenisData } = req.body;
        const filePath = req.file.path;

        // Membaca file Excel menggunakan SheetJS
        const workbook = XLSX.readFile(filePath, { defval: "" });
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
        
        // Hapus file fisik sementara di folder uploads/
        fs.unlinkSync(filePath);

        // Petakan metadata ke setiap baris
        const newRows = sheetData.map(row => ({
            Bulan: bulan || '-',
            Gelombang: gelombang || '-',
            JenisData: jenisData || 'Lainnya',
            GolonganJenis: golongan || 'Lainnya',
            ...row
        }));

        let allData = getSavedData();
        
        // Logika Overwrite / Replace Data Lama jika Kombinasi Unik Sama
        allData = allData.filter(item => 
            !(item.Bulan === bulan && 
              item.Gelombang === gelombang && 
              (item.JenisData || 'Lainnya') === (jenisData || 'Lainnya') && 
              (item.GolonganJenis || 'Lainnya') === (golongan || 'Lainnya'))
        );

        // Gabungkan dan simpan
        allData = allData.concat(newRows);
        saveAllData(allData);

        res.json({
            success: true,
            message: 'Data berhasil diarsipkan ke server!',
            totalData: allData.length,
            data: allData
        });

    } catch (error) {
        console.error('Error saat upload:', error);
        res.status(500).json({ success: false, message: 'Gagal memproses berkas Excel.' });
    }
});

// 3. ENDPOINT HAPUS DATA (UNTUK TOMBOL MARK & MARK ALL)
app.post('/api/delete', (req, res) => {
    try {
        const { indices } = req.body;
        let allData = getSavedData();

        if (indices && Array.isArray(indices)) {
            // Urutkan indeks secara descending agar saat splice indeks tidak bergeser
            const sortedIndices = [...indices].sort((a, b) => b - a);
            sortedIndices.forEach(idx => {
                if (idx >= 0 && idx < allData.length) {
                    allData.splice(idx, 1);
                }
            });

            saveAllData(allData);
            return res.json({ success: true, message: `${indices.length} baris data berhasil dihapus!`, data: allData });
        } else {
            return res.status(400).json({ success: false, message: 'Indeks data tidak valid.' });
        }
    } catch (error) {
        console.error('Error saat hapus data:', error);
        res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server saat menghapus data.' });
    }
});

// Jalankan Server
app.listen(3000, () => {
    console.log('Server berjalan aktif di http://localhost:3000');
});