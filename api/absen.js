export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ message: 'Method not allowed' });
    }

    const { uid } = req.body;
    if (!uid) {
        return res.status(400).json({ message: 'UID tidak boleh kosong' });
    }

    const CLIENT_ID = process.env.AZURE_CLIENT_ID;
    const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
    const REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;
    const FILE_ID = process.env.MS_FILE_ID;

    try {
        // 1. Dapatkan Access Token baru menggunakan Refresh Token
        const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET,
                refresh_token: REFRESH_TOKEN,
                grant_type: 'refresh_token',
                scope: 'Files.ReadWrite offline_access'
            })
        });

        const tokenData = await tokenRes.json();
        if (!tokenRes.ok) throw new Error(tokenData.error_description || 'Gagal memperbarui token Microsoft');

        const accessToken = tokenData.access_token;

        // 2. Baca Sheet 'Data_Karyawan' untuk mencocokkan UID RFID
        const getSheetRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${FILE_ID}/workbook/worksheets('Data_Karyawan')/usedRange`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const sheetData = await getSheetRes.json();
        if (!getSheetRes.ok) throw new Error('Gagal membaca data karyawan dari Excel');

        const rows = sheetData.values; // Baris data dari Excel
        // Asumsi Kolom A = UID Kartu, Kolom B = Nama Karyawan
        let namaKaryawan = null;
        for (let i = 1; i < rows.length; i++) { // Mulai dari baris ke-2 (lewati header)
            if (String(rows[i][0]).trim() === String(uid).trim()) {
                namaKaryawan = rows[i][1];
                break;
            }
        }

        if (!namaKaryawan) {
            return res.status(404).json({ message: 'Kartu RFID tidak terdaftar di sistem!' });
        }

        // 3. Catat ke Sheet 'Log_Absensi'
        const waktuSekarang = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
        const addRowRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${FILE_ID}/workbook/worksheets('Log_Absensi')/tables/TabelLog/rows`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                values: [[uid, namaKaryawan, waktuSekarang, 'Hadir']]
            })
        });

        if (!addRowRes.ok) throw new Error('Gagal mencatat log absensi ke Excel');

        return res.status(200).json({ success: true, nama: namaKaryawan, waktu: waktuSekarang });

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: error.message });
    }
}