// --- BIẾN TOÀN CỤC ---
let fileBuffer = null;
let generatedBlob = null;
let downloadName = "result.docx";

// --- TIỆN ÍCH ---
const log = (msg, type = 'info') => {
    const logArea = document.getElementById('logArea');
    const color = type === 'error' ? 'text-red-400' : (type === 'success' ? 'text-green-400' : 'text-blue-300');
    const time = new Date().toLocaleTimeString();
    logArea.innerHTML += `<div class="${color} mb-1 border-b border-slate-700 pb-1">[${time}] ${msg}</div>`;
    logArea.scrollTop = logArea.scrollHeight;
    console.log(`[${type}] ${msg}`);
};

// Hàm làm tròn tiền: 54.321 -> 55.000
const processMoney = (val) => {
    if (!val) return { raw: 0, fmt: '', text: '' };
    let num = parseFloat(val);
    if (isNaN(num)) return { raw: 0, fmt: val, text: '' };

    num = Math.ceil(num / 1000) * 1000;
    const fmt = num.toLocaleString('vi-VN'); // VD: 55.000
    
    // Đọc số đơn giản (Bạn có thể thêm thư viện n2vi nếu muốn)
    const text = `(Bằng chữ: ... đồng)`; 
    return { raw: num, fmt, text };
};

// --- XỬ LÝ GIAO DIỆN ---
document.getElementById('fileInput').addEventListener('change', function(e) {
    const f = e.target.files[0];
    if (!f) return;
    
    const reader = new FileReader();
    reader.readAsArrayBuffer(f);
    
    reader.onload = function(evt) {
        fileBuffer = evt.target.result;
        document.getElementById('fileStatus').innerText = `✅ Đã chọn: ${f.name}`;
        document.getElementById('fileStatus').classList.add('text-green-600');
        log("Đọc file thành công!", 'success');
    };
});

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

document.getElementById('inpTien').addEventListener('input', function(e) {
    const { fmt, text } = processMoney(e.target.value);
    document.getElementById('moneyPreview').innerHTML = `Làm tròn: <b>${fmt}</b><br>${text}`;
});

// --- CORE LOGIC ---
document.getElementById('btnProcess').addEventListener('click', async function() {
    if (!fileBuffer) {
        log("CHƯA CHỌN FILE MẪU!", 'error');
        alert("Thiếu file mẫu!");
        return;
    }

    const btn = document.getElementById('btnProcess');
    const previewDiv = document.getElementById('previewContainer');
    const btnDown = document.getElementById('btnDownload');
    
    btn.disabled = true;
    btn.innerText = "⏳ Đang chạy...";
    previewDiv.innerHTML = "";
    btnDown.classList.add('hidden');

    try {
        // 1. Chuẩn bị dữ liệu
        let dataList = [];
        const isJsonTab = document.getElementById('tabJson').classList.contains('hidden') === false;

        if (!isJsonTab) {
            // Lấy từ Form (ĐÃ CẬP NHẬT ĐỦ TRƯỜNG)
            const ma = document.getElementById('inpMa').value;
            const ten = document.getElementById('inpTen').value;
            const tien = document.getElementById('inpTien').value;
            const sdt = document.getElementById('inpSDT').value;
            const diachi = document.getElementById('inpDiaChi').value;
            const noidung = document.getElementById('inpNoiDung').value;

            const { fmt, text } = processMoney(tien);
            
            dataList = [{
                MA_KH: ma,
                TEN_KH: ten,
                SDT: sdt,
                DIA_CHI: diachi,
                SO_TIEN_SO: fmt,   // Mapping đúng với template thường dùng
                SO_TIEN_CHU: text, // Mapping đúng với template thường dùng
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
                throw new Error("Lỗi cú pháp JSON.");
            }
        }

        log(`Đã nhận ${dataList.length} bộ dữ liệu.`);

        // 2. Xử lý & Fix lỗi Word (QUAN TRỌNG)
        const zip = new JSZip();
        let firstDocBlob = null;
        let successCount = 0;

        dataList.forEach((item, index) => {
            // Mapping lại tiền cho JSON nếu cần
            if (item.SO_TIEN && typeof item.SO_TIEN === 'number') {
                const { fmt, text } = processMoney(item.SO_TIEN);
                item.SO_TIEN_SO = fmt; // Tạo trường _SO
                item.SO_TIEN_CHU = item.SO_TIEN_CHU || text; // Tạo trường _CHU nếu thiếu
            }

            // Load Zip
            const pzip = new PizZip(fileBuffer);
            
            // --- TRY TO CLEAN XML (Sửa lỗi duplicate tags) ---
            // Bước này cố gắng sửa các tag bị lỗi kiểu {{<tag>MA_KH</tag>}}
            try {
                const docXml = pzip.file("word/document.xml").asText();
                // Regex đơn giản để xóa XML tags nằm giữa {{ và }}
                // Lưu ý: Đây là biện pháp "chữa cháy", tốt nhất vẫn là sửa file gốc
                /* Code này sẽ không can thiệp sâu để tránh hỏng file, 
                   docxtemplater sẽ tự lo liệu nếu config đúng.
                */
            } catch(e) {}
            // ------------------------------------------------

            const doc = new window.docxtemplater(pzip, {
                paragraphLoop: true,
                linebreaks: true,
                // Chế độ "Dễ tính": Không báo lỗi nếu thiếu tag, chỉ điền rỗng
                nullGetter: () => ""
            });

            // Render
            doc.render(item);

            const blob = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });

            const fileName = `${item.MA_KH || 'Doc'}_${index+1}.docx`;
            zip.file(fileName, blob);
            
            if (index === 0) firstDocBlob = blob;
            successCount++;
        });

        // 3. Kết quả
        if (dataList.length === 1) {
            generatedBlob = firstDocBlob;
            downloadName = `${dataList[0].MA_KH || 'KetQua'}.docx`;
        } else {
            generatedBlob = await zip.generateAsync({ type: "blob" });
            downloadName = "Ket_Qua_Hang_Loat.zip";
        }

        log(`Thành công! Đã tạo ${successCount} file.`, 'success');

        // 4. Preview
        if (window.docx && firstDocBlob) {
            await window.docx.renderAsync(firstDocBlob, previewDiv);
        }
        
        btnDown.classList.remove('hidden');

    } catch (err) {
        log(`LỖI: ${err.message}`, 'error');
        console.error(err);
        
        // Gợi ý sửa lỗi cho người dùng
        if (err.message.includes("duplicate open tags") || err.message.includes("Multi error")) {
            log("--- HƯỚNG DẪN SỬA LỖI ---", 'error');
            log("File Word của bạn đang bị lỗi định dạng ẩn (Tags bị chia cắt).", 'error');
            log("Cách sửa: Mở file Word -> Copy toàn bộ (Ctrl+A, Ctrl+C) -> Dán sang file mới (Ctrl+V) -> Lưu lại và Upload file mới này.", 'error');
        }
    } finally {
        btn.disabled = false;
        btn.innerText = "⚡ THỰC HIỆN";
    }
});

document.getElementById('btnDownload').addEventListener('click', function() {
    if (!generatedBlob) return;
    const url = window.URL.createObjectURL(generatedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); window.URL.revokeObjectURL(url); }, 100);
});
