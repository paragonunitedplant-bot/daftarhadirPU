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

        const rows = sheetData.values; 
        let namaKaryawan = null;
        let timKaryawan = '-';
        
        // Asumsi Kolom A = UID, Kolom B = Nama, Kolom C = Tim
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]).trim() === String(uid).trim()) {
                namaKaryawan = rows[i][1];
                timKaryawan = rows[i][2] || '-';
                break;
            }
        }

        if (!namaKaryawan) {
            return res.status(404).json({ message: 'Kartu RFID tidak terdaftar di sistem!' });
        }

        // 3. Catat ke Sheet 'Log_Absensi' (Menyesuaikan 5 Kolom: Tanggal, Jam, UID, Nama, Tim)
        const now = new Date();
        const tanggalStr = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta' });
        const jamStr = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta' });

        const addRowRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${FILE_ID}/workbook/worksheets('Log_Absensi')/tables/TabelLog/rows`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                values: [[tanggalStr, jamStr, uid, namaKaryawan, timKaryawan]]
            })
        });

        if (!addRowRes.ok) throw new Error('Gagal mencatat log absensi ke Excel');

        return res.status(200).json({ success: true, nama: namaKaryawan, waktu: `${tanggalStr} ${jamStr}` });

    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: error.message });
    }
}
