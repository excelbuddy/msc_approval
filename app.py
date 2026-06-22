import streamlit as st
import requests
import pandas as pd
from datetime import datetime
import os
import io
import urllib3
import unicodedata
import re
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from requests.adapters import HTTPAdapter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ====== SSL FIX: bypass DH_KEY_TOO_SMALL cho muasamcong.mpi.gov.vn ======
class LegacySSLAdapter(HTTPAdapter):
    """Cho phép kết nối tới server dùng DH key yếu (legacy SSL)."""
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.set_ciphers("DEFAULT:@SECLEVEL=1")
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        kwargs['ssl_context'] = ctx
        return super().init_poolmanager(*args, **kwargs)

def make_session():
    s = requests.Session()
    s.mount("https://", LegacySSLAdapter())
    return s

# ====== CẤU HÌNH EMAIL ======
try:
    SENDER_EMAIL    = st.secrets["SENDER_EMAIL"]
    SENDER_APP_PASS = st.secrets["SENDER_APP_PASS"]
except KeyError:
    st.error("⚠️ Chưa cấu hình Secrets! Vào Settings → Secrets để thêm SENDER_EMAIL và SENDER_APP_PASS.")
    st.stop()

MODE_SEPARATE   = "separate"
MODE_MULTISHEET = "multisheet"
MODE_ONESHEET   = "onesheet"

EMAIL_RECIPIENTS = [
    ("beatme", "duylinhvinhphuc@gmail.com"),
    ("beatme", "duylinh93@gmail.com"),
    ("Bui Viet Huy", "bvhuy.ho@vietcombank.com.vn"),
    ("Can Quang Minh", "minhcq.ho@vietcombank.com.vn"),
    ("Do Thi Vui", "VUIDT.HO@vietcombank.com.vn"),
    ("Do Thuy Linh", "linhdt1.ho@vietcombank.com.vn"),
    ("Khieu Van Truong", "TRUONGKV.HO@vietcombank.com.vn"),
    ("Luong Trung Hieu", "HIEULT1.HO@vietcombank.com.vn"),
    ("MAI XUAN LICH", "lichmx.ho@vietcombank.com.vn"),
    ("Nguyen Anh Thu", "nathu.ho@vietcombank.com.vn"),
    ("Nguyen Duc Huy", "HUYND.HO@vietcombank.com.vn"),
    ("Nguyen Hoai", "hoain.ho@vietcombank.com.vn"),
    ("Nguyen Kim Nhung", "NHUNGNK.HO@vietcombank.com.vn"),
    ("Nguyen Ngoc Chi Linh", "LINHNNC.HO@vietcombank.com.vn"),
    ("Nguyen Ngoc Khanh", "khanhnn.ho@vietcombank.com.vn"),
    ("Nguyen The Duy", "duynt.ho@vietcombank.com.vn"),
    ("Nguyen Thi Ngoc Bich", "ntnbich.ho2@vietcombank.com.vn"),
    ("Nguyen Thi Ngoc Mai", "MAINTN1.HO@vietcombank.com.vn"),
    ("NGUYEN THU HUONG", "nthuong.ho2@vietcombank.com.vn"),
    ("Nguyen Tu Anh", "NTANH1.HO@vietcombank.com.vn"),
    ("Nguyen Viet Nga", "nganv.ho@vietcombank.com.vn"),
    ("Nguyen Viet Tung", "TUNGNV1.HO@vietcombank.com.vn"),
    ("Pham Kim Ngan", "nganpk.ho@vietcombank.com.vn"),
    ("Pham Thi Thanh Nga", "ngaptt.ho@vietcombank.com.vn"),
    ("Phan Thi Phuong", "phuongpt1.ho@vietcombank.com.vn"),
    ("Phan Thuy Linh", "Linhpt.ho@vietcombank.com.vn"),
    ("Tran Hoang Nga", "ngath.ho1@vietcombank.com.vn"),
    ("Tran Manh Hung", "HUNGTM4.HO@vietcombank.com.vn"),
    ("Tran Quang Anh", "ANHTQ.HO@vietcombank.com.vn"),
    ("Tran Thi Cam Van", "VANTTC.HO@vietcombank.com.vn"),
    ("Tran Thi Huyen Trang", "trangtth.ho2@vietcombank.com.vn"),
    ("Tran Thi Mai Huong", "ttmhuong.ho@vietcombank.com.vn"),
    ("Tran Thi Thanh Hao", "haottt.ho@vietcombank.com.vn"),
    ("Truong Duc Hai", "haitd.ho@vietcombank.com.vn"),
    ("Vu Duy Linh", "LINHVD.HO@vietcombank.com.vn"),
    ("Vu Thi Bich", "BICHVT.HO@vietcombank.com.vn"),
    ("Vu Thu Nga", "ngavt.ho@vietcombank.com.vn"),
    ("Vu Viet Anh", "anhvv1.ho@vietcombank.com.vn"),
]

# ====== TIỆN ÍCH ======
def remove_accents(s):
    nfkd = unicodedata.normalize('NFKD', s)
    s2   = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'[^a-zA-Z0-9_]+', '', s2.replace(' ', '_')).lower()

# ====== XỬ LÝ DỮ LIỆU ======
def process_data_for_excel(response_data):
    if not response_data or 'page' not in response_data or 'content' not in response_data['page']:
        return None
    processed = []
    for item in response_data['page']['content']:
        row = {k: item.get(k, '') for k in item}
        for k in item:
            if isinstance(item[k], list):
                row[k] = '; '.join(str(x) for x in item[k] if x is not None)
        if 'locations' in item and item['locations']:
            row['locations'] = '; '.join(
                f"{loc.get('provName','')} ({loc.get('provCode','')})"
                for loc in item['locations']
            )
        processed.append(row)
    return processed

def _clean_df(data, keyword=None):
    df = pd.DataFrame(data)
    drop_cols = ["id","type","tab","soQuyetDinh","ngayBanHanhQuyetDinh","decisions"]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors='ignore')
    if keyword is not None:
        df.insert(0, 'tu_khoa', keyword)
    return df

def _auto_width(ws):
    for col in ws.columns:
        w, letter = 0, col[0].column_letter
        for cell in col:
            try: w = max(w, len(str(cell.value)))
            except: pass
        ws.column_dimensions[letter].width = min(w + 2, 50)

# ====== FETCH ======
def fetch_keyword_raw(keyword, total_pages, log_func=None):
    url = "https://muasamcong.mpi.gov.vn/o/egp-portal-personal-page/services/smart/search_prc"
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0'
    }
    cookies = {'GUEST_LANGUAGE_ID': 'vi_VN', 'COOKIE_SUPPORT': 'true'}
    payload = [{"pageSize": 50, "pageNumber": 0, "query": [{
        "index": "es-smart-pricing", "keyWord": keyword, "matchType": "all-1",
        "matchFields": ["danh_muc_hang_hoa","ma_hs","xuat_xu","ma_tbmt",
                        "ky_ma_hieu","nhan_hieu","hang_san_xuat"],
        "filters": [
            {"fieldName": "type", "searchType": "in", "fieldValues": ["HANG_HOA"]},
            {"fieldName": "tab",  "searchType": "in", "fieldValues": ["HANG_HOA"]}
        ]
    }]}]

    def log(m):
        if log_func: log_func(m)

    all_data = []
    session  = make_session()
    for page in range(total_pages):
        log(f"  📄 [{keyword}] Trang {page+1}/{total_pages}...")
        payload[0]["pageNumber"] = page
        try:
            r = session.post(url, headers=headers, cookies=cookies,
                             json=payload, verify=False, timeout=30)
            if r.status_code == 200:
                d = process_data_for_excel(r.json())
                if not d:
                    log(f"  ⚠️  [{keyword}] Không có dữ liệu trang {page+1} → Dừng quét")
                    break
                all_data.extend(d)
                if len(d) < 50:
                    log(f"  ✅ [{keyword}] Đã tới trang cuối ({len(d)} bản ghi)")
                    break
            else:
                log(f"  ❌ [{keyword}] HTTP {r.status_code}")
        except Exception as e:
            log(f"  ❌ [{keyword}] Lỗi: {e}")
    return all_data

# ====== GỬI EMAIL ======
def send_email_to(recipient_name, recipient_email, results_summary, excel_files, save_mode):
    try:
        msg = MIMEMultipart()
        msg['From']    = SENDER_EMAIL
        msg['To']      = recipient_email
        msg['Subject'] = f"[Email tự động] Kết quả tra cứu dữ liệu hàng hóa muasamcong - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

        lines = [
            f"Xin chào {recipient_name},",
            "\n\nLƯU Ý KHI SỬ DỤNG DỮ LIỆU:\n",
            " - Nên dùng các dòng dữ liệu mà cột bidForm có giá trị [DTRR], [CHCT].",
            " - Chọn mức giá hợp lý để đảm bảo tính khả thi của gói mua sắm.",
            " - Dữ liệu chỉ dùng nội bộ trong Ban mua sắm, không gửi ra nơi khác!",
            "",
            "Kết quả tìm kiếm hàng hóa trên muasamcong:", "",
        ]
        for r in results_summary:
            lines.append(f"  • \"{r['keyword']}\" → {r['count']} bản ghi")
        lines += ["", f"File đính kèm: {len(excel_files)} file.", "", "© Developed by Beatme!"]

        msg.attach(MIMEText('\n'.join(lines), 'plain', 'utf-8'))

        for fp in excel_files:
            if not os.path.exists(fp): continue
            with open(fp, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f'attachment; filename="{os.path.basename(fp)}"')
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(SENDER_EMAIL, SENDER_APP_PASS)
            s.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        return True
    except Exception as e:
        return False

# ====== LƯU EXCEL VÀO BUFFER ======
def build_excel_buffer(save_mode_val, all_kw_data):
    buffer = io.BytesIO()
    if save_mode_val == MODE_ONESHEET:
        frames = [_clean_df(data, kw) for kw, data in all_kw_data if data]
        if not frames: return None
        df = pd.concat(frames, ignore_index=True)
        with pd.ExcelWriter(buffer, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='msc_data', index=False)
            _auto_width(w.sheets['msc_data'])

    elif save_mode_val == MODE_MULTISHEET:
        with pd.ExcelWriter(buffer, engine='openpyxl') as w:
            for kw, data in all_kw_data:
                if not data: continue
                df = _clean_df(data)
                sn = base = remove_accents(kw)[:28] or "sheet"
                i = 2
                while sn in w.sheets: sn = f"{base[:25]}_{i}"; i += 1
                df.to_excel(w, sheet_name=sn, index=False)
                _auto_width(w.sheets[sn])

    elif save_mode_val == MODE_SEPARATE:
        # Với chế độ separate, gộp tất cả vào 1 buffer (mỗi sheet = 1 từ khóa)
        with pd.ExcelWriter(buffer, engine='openpyxl') as w:
            for kw, data in all_kw_data:
                if not data: continue
                df = _clean_df(data)
                sn = base = remove_accents(kw)[:28] or "sheet"
                i = 2
                while sn in w.sheets: sn = f"{base[:25]}_{i}"; i += 1
                df.to_excel(w, sheet_name=sn, index=False)
                _auto_width(w.sheets[sn])

    buffer.seek(0)
    return buffer

# ======================================================================
# GIAO DIỆN STREAMLIT
# ======================================================================
st.set_page_config(page_title="Dữ liệu muasamcong", page_icon="🔍", layout="centered")
st.title("🔍 Tra cứu dữ liệu Mua Sắm Công")

# ── Từ khóa ────────────────────────────────────────────────────────────
keyword_input = st.text_input(
    "Từ khóa hàng hóa (nhiều từ khóa ngăn cách bằng dấu ';')",
    value="iphone; máy chủ; laptop"
)

# ── Số trang ───────────────────────────────────────────────────────────
total_pages = st.number_input(
    "Số trang / từ khóa (50 bản ghi/trang)",
    min_value=1, max_value=500, value=100
)

# ── Chế độ lưu file ────────────────────────────────────────────────────
save_mode_label = st.radio("Chế độ lưu file", [
    "Gộp vào 1 file – tất cả vào 1 sheet (có cột 'tu_khoa')",
    "Mỗi từ khóa → 1 file Excel riêng (tải về từng file)",
    "Gộp vào 1 file – mỗi từ khóa = 1 sheet riêng",
])
save_mode_val = {
    "Gộp vào 1 file – tất cả vào 1 sheet (có cột 'tu_khoa')": MODE_ONESHEET,
    "Mỗi từ khóa → 1 file Excel riêng (tải về từng file)": MODE_SEPARATE,
    "Gộp vào 1 file – mỗi từ khóa = 1 sheet riêng": MODE_MULTISHEET,
}[save_mode_label]

# ── Người nhận email ───────────────────────────────────────────────────
st.markdown("**📧 Gửi email tới:**")

recip_options = [f"{name.strip()} <{email}>" for name, email in EMAIL_RECIPIENTS]
selected_recip = st.multiselect(
    "Chọn từ danh sách:",
    options=recip_options,
    placeholder="Tìm tên hoặc email..."
)

# Nhập email tùy chỉnh
custom_email_input = st.text_input(
    "✏️ Hoặc nhập thêm email khác (nhiều email ngăn cách bằng dấu ';'):",
    placeholder="vd: nguyen@example.com; tran@example.com"
)

# Tổng hợp danh sách người nhận
selected_recipients = [
    EMAIL_RECIPIENTS[recip_options.index(s)] for s in selected_recip
]
if custom_email_input.strip():
    for raw in custom_email_input.split(';'):
        raw = raw.strip()
        if raw and '@' in raw:
            selected_recipients.append(("(Tùy chỉnh)", raw))

if selected_recipients:
    st.caption(f"Sẽ gửi tới {len(selected_recipients)} địa chỉ: " +
               ", ".join(e for _, e in selected_recipients))
else:
    st.caption("Không chọn ai → chỉ tải file, không gửi email.")

st.divider()

# ── Nút chạy ───────────────────────────────────────────────────────────
if st.button("🚀 LẤY DỮ LIỆU", type="primary", use_container_width=True):
    keywords = [k.strip() for k in keyword_input.split(';') if k.strip()]
    if not keywords:
        st.error("Vui lòng nhập ít nhất 1 từ khóa!")
        st.stop()

    log_area   = st.empty()
    log_lines  = []

    def log(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines), language=None)

    results_summary, all_kw_data = [], []

    for kw in keywords:
        log(f"🔍 Tìm kiếm: \"{kw}\"")
        data = fetch_keyword_raw(kw, int(total_pages), log)
        if data:
            log(f"  ✅ \"{kw}\": {len(data)} bản ghi\n")
            all_kw_data.append((kw, data))
            results_summary.append({"keyword": kw, "count": len(data)})
        else:
            log(f"  ⚠️  \"{kw}\": không có kết quả\n")
            results_summary.append({"keyword": kw, "count": 0})

    if not all_kw_data:
        st.warning("Không tìm thấy dữ liệu nào.")
        st.stop()

    # Lưu file Excel
    log("💾 Đang tạo file Excel...")
    fname  = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    buffer = build_excel_buffer(save_mode_val, all_kw_data)

    if buffer:
        st.success("✅ Hoàn thành! Nhấn nút bên dưới để tải file.")
        st.download_button(
            label="📥 Tải file Excel",
            data=buffer,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Gửi email nếu có người nhận
        if selected_recipients:
            log(f"\n📧 Đang gửi email tới {len(selected_recipients)} người...")
            tmp_path = f"/tmp/{fname}"
            with open(tmp_path, 'wb') as f:
                f.write(buffer.getvalue())
            for name, email_addr in selected_recipients:
                log(f"  → {name} <{email_addr}>")
                ok = send_email_to(name, email_addr, results_summary, [tmp_path], save_mode_val)
                log("    ✅ Gửi thành công!" if ok else "    ❌ Gửi thất bại!")
        else:
            log("\nℹ️  Không có người nhận → bỏ qua gửi email.")
    else:
        st.error("Không thể tạo file Excel.")
