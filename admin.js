import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import {
    getFirestore, collection, query, where, getDocs, doc,
    setDoc,
    serverTimestamp,
    orderBy
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyA47cTihvHBnqYcn4HDlOKJf88O7MmzINo",
    authDomain: "absenit-92dd9.firebaseapp.com",
    projectId: "absenit-92dd9",
    storageBucket: "absenit-92dd9.firebasestorage.app",
    messagingSenderId: "219381976986",
    appId: "1:219381976986:web:7f1ac4bd70470a52e63423"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

document.getElementById('exportAdminBtn').addEventListener('click', exportDataKeExcel);
document.getElementById('setTokenBtn').addEventListener('click', updateToken);

async function updateToken() {
    const token = document.getElementById('tokenInput').value.trim();
    const duration = parseInt(document.getElementById('durationInput').value);

    if (!token || isNaN(duration)) {
        return alert("Harap isi kode token dan durasi dengan benar!");
    }

    try {
        // Simpan ke dokumen khusus 'token_aktif' di koleksi 'system'
        await setDoc(doc(db, "system", "token_aktif"), {
            token: token,
            createdAt: serverTimestamp(), // Waktu server agar akurat
            duration: duration
        });
        alert("Token '" + token + "' aktif selama " + duration + " menit!");
    } catch (e) {
        console.error("Gagal update token: ", e);
        alert("Gagal memperbarui token.");
    }
}

async function exportDataKeExcel() {
    const tanggalInput = document.getElementById('filterDate').value;
    if (!tanggalInput) return alert("Pilih tanggal dulu!");

    const start = new Date(tanggalInput);
    start.setHours(0, 0, 0, 0);
    const end = new Date(tanggalInput);
    end.setHours(23, 59, 59, 999);

    // --- UPDATE: Tambahkan orderBy timestamp agar data diurutkan dari yang terlama ke terbaru ---
    const q = query(
        collection(db, "presensi"),
        where("timestamp", ">=", start),
        where("timestamp", "<=", end),
        orderBy("timestamp", "asc") // Urutkan biar yang baru diproses terakhir
    );

    const querySnapshot = await getDocs(q);
    const dataFirebase = {};

    querySnapshot.forEach((doc) => {
        const d = doc.data();
        const key = `${d.nama}-${d.kelas}`;

        // Karena orderBy asc, jika ada nama yang sama, 
        // d.ttd yang baru akan menimpa (overwrite) yang lama di dalam object.
        dataFirebase[key] = d.ttd;
    });

    // 2. Load File Excel
    const response = await fetch('assets/Absensi.xlsx');
    const arrayBuffer = await response.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    let worksheet = workbook.getWorksheet('JULI');
    if (!worksheet) {
        worksheet = workbook.worksheets[workbook.worksheets.length - 1];
    }

    const getVal = (row, colIndex) => {
        const cell = row.getCell(colIndex);
        if (!cell || cell.value === null || cell.value === undefined) return '';
        if (typeof cell.value === 'object') {
            if (cell.value.result !== undefined) return String(cell.value.result).trim();
            if (cell.value.richText) return cell.value.richText.map(t => t.text).join('').trim();
            if (cell.value.text !== undefined) return String(cell.value.text).trim();
        }
        return String(cell.text || cell.value || '').trim();
    };

    worksheet.eachRow(async (row, rowNumber) => {
        if (rowNumber >= 8) {
            const noStr = getVal(row, 1);
            const nama = getVal(row, 2);
            const kelas = getVal(row, 3);

            if (/^\d+$/.test(noStr) && nama !== "") {
                const no = parseInt(noStr, 10);
                const key = `${nama}-${kelas}`;

                if (dataFirebase[key]) {
                    // Atur tinggi baris agar tidak terlalu sesak untuk ukuran 20x20
                    row.height = 30;

                    const base64Data = dataFirebase[key];
                    const imageId = workbook.addImage({
                        base64: base64Data,
                        extension: 'png',
                    });

                    const colTarget = (no % 2 !== 0) ? 4 : 5;
                    const cell = row.getCell(colTarget);

                    cell.value = {
                        richText: [
                            { text: no + '. ', font: { bold: true, size: 12 } },
                        ]
                    };

                    cell.alignment = {
                        vertical: 'middle',
                        horizontal: 'left', // Rata kiri sesuai permintaan
                        indent: 1 // Memberi sedikit ruang agar tidak nempel garis cell
                    };

                    worksheet.addImage(imageId, {
                        tl: {
                            col: colTarget - 0.3,
                            row: rowNumber - 0.9
                        },
                        ext: { width: 50, height: 30 },
                        editAs: 'oneCell'
                    });

                    // Bersihkan kolom zig-zag pasangannya
                    const otherCol = (colTarget === 4) ? 5 : 4;
                    row.getCell(otherCol).value = "";
                }
            }
        }
    });

    // 4. Download Hasilnya
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Rekap_Absen_${tanggalInput}.xlsx`;
    a.click();
}