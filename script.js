// --- BIẾN TOÀN CỤC ---
let fileBuffer = null;
let generatedBlob = null;
let downloadName = "result.docx";

// --- TIỆN ÍCH ---
const log = (msg, type = 'info') => {
    const logArea = document.getElementById('logArea');
    const color = type === 'error' ? 'text-red-400' : (type === 'success' ? 'text-green-400' : 'text-blue-300');
    const time = new Date().toLocaleTimeString();
    logArea.innerHTML += `<div class="${color}">[${time}] ${msg}</div>`;
    logArea.scrollTop = logArea.scrollHeight;
    console.log(`[${type}] ${msg}`);
};

// Hàm làm tròn tiền: 54.321 -> 55.000
const processMoney = (val) => {
    if (!val) return { raw: 0, fmt: '', text: '' };
    let num = parseFloat(val);
    if (isNaN(num)) return { raw: 0, fmt: val, text: '' };

    num = Math.ceil(num / 1000) * 1000;
    const fmt = num.toLocaleString('vi-VN');
    const text = `(Bằng chữ: ... đồng)`; // Có thể thêm thư viện đọc số nếu cần
    return { raw: num, fmt, text };
};

// --- XỬ LÝ GIAO DIỆN ---
// 1. Chọn file
document.getElementById('fileInput').addEventListener('change', function(e) {
    const f = e.target.files[0];
    if (!f) return;
    
    log(`Đang đọc file: ${f.name}...`);
    const reader = new FileReader();
    reader.readAsArrayBuffer(f);
    
    reader.onload = function(evt) {
        fileBuffer = evt.target.result;
        document.getElementById('fileStatus').innerText = `✅ Đã chọn: ${f.name}`;
        document.getElementById('fileStatus').classList.add('text-green-600');
        log("Đọc file thành công!", 'success');
    };
    
    reader.onerror = function() {
        log("Lỗi không đọc được file!", 'error');
    };
});

// 2. Chuyển Tab
window.switchTab = (tabName) => {
    const tabForm = document.getElementById('tabForm');
    const tabJson = document.getElementById('tabJson');
    const btns = document.querySelectorAll('.tab-btn');

    if (tabName === 'form') {
        tabForm.classList.remove('hidden');
        tabJson.classList.add('hidden');
        btns[0].classList.add('active');
        btns[1].classList.remove('active');
    } else {
        tabForm.classList.add('hidden');
        tabJson.classList.remove('hidden');
        btns[0].classList.remove('active');
        btns[1].classList.add('active');
    }
};

// 3. Tự động cập nhật preview tiền
document.getElementById('inpTien').addEventListener('input', function(e) {
    const { fmt, text } = processMoney(e.target.value);
    document.getElementById('moneyPreview').innerHTML = `Làm tròn: <b>${fmt}</b><br>${text}`;
});

// --- CORE LOGIC (QUAN TRỌNG) ---
document.getElementById('btnProcess').addEventListener('click', async function() {
    if (!fileBuffer) {
        log("CHƯA CHỌN FILE MẪU! Hãy upload file .docx trước.", 'error');
        alert("Thiếu file mẫu!");
        return;
    }

    log("Bắt đầu xử lý...");
    const btn = document.getElementById('btnProcess');
    const previewDiv = document.getElementById('previewContainer');
    const btnDown = document.getElementById('btnDownload');
    
    btn.disabled = true;
    btn.innerText = "⏳ Đang chạy...";
    previewDiv.innerHTML = ""; // Xóa preview cũ
    btnDown.classList.add('hidden');

    try {
        // A. Lấy dữ liệu
        let dataList = [];
        const isJsonTab = document.getElementById('tabJson').classList.contains('hidden') === false;

        if (!isJsonTab) {
            // Lấy từ Form
            const ma = document.getElementById('inpMa').value;
            const ten = document.getElementById('inpTen').value;
            const tien = document.getElementById('inpTien').value;
            const noidung = document.getElementById('inpNoiDung').value;

            const { fmt, text } = processMoney(tien);
            
            dataList = [{
                MA_KH: ma,
                TEN_KH: ten,
                SO_TIEN: fmt,
                SO_TIEN_CHU: text,
                NOI_DUNG: noidung
            }];
        } else {
            // Lấy từ JSON
            const jsonVal = document.getElementById('inpJson').value;
            if (!jsonVal.trim()) throw new Error("Ô JSON đang trống!");
            
            try {
                dataList = JSON.parse(jsonVal);
                if (!Array.isArray(dataList)) dataList = [dataList];
            } catch (e) {
                throw new Error("Lỗi cú pháp JSON: " + e.message);
            }
        }

        log(`Đã nhận ${dataList.length} bộ dữ liệu.`);

        // B. Xử lý Docxtemplater & Zip
        const zip = new JSZip();
        let firstDocBlob = null;
        let successCount = 0;

        // Vòng lặp xử lý từng item
        dataList.forEach((item, index) => {
            // Format tiền nếu trong JSON là số nguyên
            if (item.SO_TIEN && typeof item.SO_TIEN === 'number') {
                const { fmt, text } = processMoney(item.SO_TIEN);
                item.SO_TIEN = fmt;
                if (!item.SO_TIEN_CHU) item.SO_TIEN_CHU = text;
            }

            // Init thư viện Word
            const pzip = new PizZip(fileBuffer);
            const doc = new window.docxtemplater(pzip, {
                paragraphLoop: true,
                linebreaks: true,
                // QUAN TRỌNG: Không báo lỗi nếu thiếu tag
                nullGetter: function(part) {
                    if (part.module === "template") return "";
                    return ""; 
                }
            });

            // Render
            doc.render(item);

            // Tạo blob file con
            const blob = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });

            const fileName = `${item.MA_KH || 'Doc'}_${index+1}.docx`;
            zip.file(fileName, blob);
            
            if (index === 0) firstDocBlob = blob;
            successCount++;
        });

        // C. Kết quả đầu ra
        if (dataList.length === 1) {
            generatedBlob = firstDocBlob;
            downloadName = `${dataList[0].MA_KH || 'KetQua'}.docx`;
        } else {
            generatedBlob = await zip.generateAsync({ type: "blob" });
            downloadName = "Ket_Qua_Hang_Loat.zip";
        }

        log(`Xử lý thành công ${successCount} file.`, 'success');

        // D. Preview file đầu tiên
        if (window.docx && firstDocBlob) {
            log("Đang render preview...");
            await window.docx.renderAsync(firstDocBlob, previewDiv);
        }

        // Hiện nút tải
        btnDown.classList.remove('hidden');

    } catch (err) {
        log(`LỖI NGHIÊM TRỌNG: ${err.message}`, 'error');
        console.error(err);
        if (err.properties && err.properties.errors) {
            err.properties.errors.forEach(e => log(`Chi tiết Word: ${e.properties.explanation}`, 'error'));
        }
    } finally {
        btn.disabled = false;
        btn.innerText = "⚡ THỰC HIỆN";
    }
});

// 4. Tải xuống (Force Download)
document.getElementById('btnDownload').addEventListener('click', function() {
    if (!generatedBlob) return;
    log(`Đang tải xuống: ${downloadName}`);
    
    const url = window.URL.createObjectURL(generatedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 100);
});

// Bắt lỗi toàn cục
window.onerror = function(msg, url, line) {
    log(`Global Error: ${msg} (Line ${line})`, 'error');
};
