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
    const fmt = num.toLocaleString('vi-VN');
    const text = `(Bằng chữ: ... đồng)`; 
    return { raw: num, fmt, text };
};

// --- HÀM VÁ LỖI FILE WORD (QUAN TRỌNG) ---
// Hàm này sẽ đi sâu vào cấu trúc XML của file Word và hàn gắn các tag bị vỡ
const patchBrokenTags = (xmlContent) => {
    // 1. Hàn gắn các thẻ {{ bị tách rời (VD: <w:t>{</w:t>...<w:t>{</w:t>)
    // Regex tìm kiếm các ký tự XML nằm giữa 2 dấu {
    let patched = xmlContent.replace(
        /(<w:t>\{<\/w:t>)([\s\S]*?)(<w:t>\{<\/w:t>)/g, 
        function(match, start, middle, end) {
            // Thay thế bằng 1 thẻ {{ duy nhất
            return `<w:t>{{</w:t>${middle}`; 
        }
    );

    // 2. Hàn gắn các thẻ }} bị tách rời
    patched = patched.replace(
        /(<w:t>\}<\/w:t>)([\s\S]*?)(<w:t>\}<\/w:t>)/g, 
        function(match, start, middle, end) {
            return `${middle}<w:t>}}</w:t>`;
        }
    );

    return patched;
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
            // Lấy từ Form
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
                SO_TIEN_SO: fmt,
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
                throw new Error("Lỗi cú pháp JSON.");
            }
        }

        log(`Đã nhận ${dataList.length} bộ dữ liệu.`);

        // 2. Xử lý & Fix lỗi Word (QUAN TRỌNG)
        const zip = new JSZip();
        let firstDocBlob = null;
        let successCount = 0;

        // ** PATCH FILE XML **
        // Mở file zip và sửa nội dung XML trực tiếp để vá lỗi
        const pzipMain = new PizZip(fileBuffer);
        const docXmlPath = "word/document.xml";
        if (pzipMain.files[docXmlPath]) {
            try {
                const originalXml = pzipMain.file(docXmlPath).asText();
                const fixedXml = patchBrokenTags(originalXml);
                pzipMain.file(docXmlPath, fixedXml); // Ghi đè file XML đã sửa vào zip
                log("Đã tự động vá lỗi tag trong file Word.", 'info');
            } catch (e) {
                console.warn("Không thể vá lỗi XML:", e);
            }
        }
        // Tạo buffer mới từ file đã vá lỗi
        const fixedBuffer = pzipMain.generate({type: "arraybuffer"});

        dataList.forEach((item, index) => {
            // Mapping lại tiền cho JSON nếu cần
            if (item.SO_TIEN && typeof item.SO_TIEN === 'number') {
                const { fmt, text } = processMoney(item.SO_TIEN);
                item.SO_TIEN_SO = fmt; 
                item.SO_TIEN_CHU = item.SO_TIEN_CHU || text;
            }

            const pzip = new PizZip(fixedBuffer);
            
            const doc = new window.docxtemplater(pzip, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter: () => ""
            });

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
        if (err.properties && err.properties.errors) {
            err.properties.errors.forEach(e => log(`Chi tiết Word: ${e.properties.explanation}`, 'error'));
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
