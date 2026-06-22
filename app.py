import streamlit as st
import requests
import pandas as pd
from datetime import datetime
import io
import json
import base64
import urllib3
import unicodedata
import re
import ssl
from requests.adapters import HTTPAdapter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ====== SSL FIX ======
class LegacySSLAdapter(HTTPAdapter):
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

# ====== CONFIG ======
try:
    TELEGRAM_BOT_TOKEN  = st.secrets["TELEGRAM_BOT_TOKEN"]
    TELEGRAM_CHAT_ID    = st.secrets["TELEGRAM_CHAT_ID"]
    GAS_WEBHOOK_URL     = st.secrets["GAS_WEBHOOK_URL"]
except KeyError as e:
    st.error(f"⚠️ Thiếu secret: {e}. Vào Settings → Secrets để bổ sung.")
    st.stop()

MODE_SEPARATE   = "separate"
MODE_MULTISHEET = "multisheet"
MODE_ONESHEET   = "onesheet"

# ====== TIỆN ÍCH ======
def remove_accents(s):
    nfkd = unicodedata.normalize('NFKD', s)
    s2   = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'[^a-zA-Z0-9_]+', '', s2.replace(' ', '_')).lower()

def is_valid_vcb_email(email: str) -> bool:
    email = email.strip().lower()
    return bool(re.match(r'^[\w.\-]+@vietcombank\.com\.vn$', email))

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

# ====== BUILD EXCEL BUFFER ======
def build_excel_buffer(save_mode_val, all_kw_data):
    buffer = io.BytesIO()
    if save_mode_val == MODE_ONESHEET:
        frames = [_clean_df(data, kw) for kw, data in all_kw_data if data]
        if not frames: return None
        df = pd.concat(frames, ignore_index=True)
        with pd.ExcelWriter(buffer, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='msc_data', index=False)
            _auto_width(w.sheets['msc_data'])

    else:  # MULTISHEET hoặc SEPARATE (đều gộp theo sheet)
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

# ====== GỬI TELEGRAM (thông báo + nút duyệt) ======
def send_telegram_approval(
    user_email: str,
    keywords: list[str],
    results_summary: list[dict],
    excel_b64: str,
    fname: str,
    save_mode_label: str,
    request_id: str,
):
    """
    Gửi message Telegram với 2 inline button: ✅ Đồng ý / ❌ Từ chối.
    Callback data chứa toàn bộ thông tin cần thiết để Apps Script xử lý.
    Vì Telegram giới hạn callback_data 64 bytes, ta lưu payload vào
    một dict tạm trong session_state (trong thực tế production nên dùng DB/Redis),
    còn callback_data chỉ chứa request_id.
    """
    total_records = sum(r['count'] for r in results_summary)
    kw_lines = "\n".join(
        f"  • \"{r['keyword']}\" -> {r['count']} ban ghi"
        for r in results_summary
    )
    # Dùng plain text để tránh lỗi Markdown parsing khi từ khóa có ký tự đặc biệt
    text = (
        f"📋 YEU CAU TRA CUU MSC\n\n"
        f"👤 Email: {user_email}\n"
        f"🕐 Thoi gian: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"📊 Che do luu: {save_mode_label}\n"
        f"📦 Tong ban ghi: {total_records}\n\n"
        f"Tu khoa:\n{kw_lines}\n\n"
        f"➡️ Nhan ✅ Dong y de gui file ket qua cho user."
    )

    # Inline keyboard
    keyboard = {
        "inline_keyboard": [[
            {"text": "✅ Đồng ý",  "callback_data": f"approve:{request_id}"},
            {"text": "❌ Từ chối", "callback_data": f"reject:{request_id}"},
        ]]
    }

    tg_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    r = requests.post(tg_url, json={
        "chat_id":      TELEGRAM_CHAT_ID,
        "text":         text,
        "reply_markup": keyboard,
    }, timeout=15)
    return r.ok, r.json()

# ====== GỬI PAYLOAD LÊN GAS WEBHOOK (kèm file base64) ======
def register_request_to_gas(
    request_id: str,
    user_email: str,
    keywords: list[str],
    results_summary: list[dict],
    excel_b64: str,
    fname: str,
    save_mode_label: str,
):
    """
    Gửi toàn bộ dữ liệu (kể cả file Excel dạng base64) lên Apps Script
    để lưu tạm, chờ callback từ Telegram Bot.
    """
    payload = {
        "action":          "register",
        "request_id":      request_id,
        "user_email":      user_email,
        "keywords":        keywords,
        "results_summary": results_summary,
        "excel_b64":       excel_b64,
        "fname":           fname,
        "save_mode":       save_mode_label,
        "submitted_at":    datetime.now().isoformat(),
    }
    r = requests.post(GAS_WEBHOOK_URL, json=payload, timeout=30)
    return r.ok

# ======================================================================
# GIAO DIỆN STREAMLIT
# ======================================================================
st.set_page_config(page_title="Dữ liệu muasamcong", page_icon="🔍", layout="centered")
st.title("🔍 Tra cứu dữ liệu Mua Sắm Công")

st.info(
    "📌 **Hướng dẫn:** Nhập thông tin, nhấn *Lấy dữ liệu*. "
    "Kết quả sẽ được gửi tới email của bạn sau khi được phê duyệt.",
    icon="ℹ️"
)

# ── Email người dùng ────────────────────────────────────────────────────
user_email = st.text_input(
    "📧 Email nhận kết quả (chỉ chấp nhận @vietcombank.com.vn)",
    placeholder="yourname@vietcombank.com.vn"
)

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
    "Gộp vào 1 file – mỗi từ khóa = 1 sheet riêng",
])
save_mode_val = {
    "Gộp vào 1 file – tất cả vào 1 sheet (có cột 'tu_khoa')": MODE_ONESHEET,
    "Gộp vào 1 file – mỗi từ khóa = 1 sheet riêng":          MODE_MULTISHEET,
}[save_mode_label]

st.divider()

# ── Nút chạy ───────────────────────────────────────────────────────────
if st.button("🚀 LẤY DỮ LIỆU VÀ GỬI YÊU CẦU PHÊ DUYỆT", type="primary", use_container_width=True):

    # Validate email
    if not user_email.strip():
        st.error("Vui lòng nhập email nhận kết quả!")
        st.stop()
    if not is_valid_vcb_email(user_email):
        st.error("❌ Chỉ chấp nhận email @vietcombank.com.vn!")
        st.stop()

    keywords = [k.strip() for k in keyword_input.split(';') if k.strip()]
    if not keywords:
        st.error("Vui lòng nhập ít nhất 1 từ khóa!")
        st.stop()

    log_area  = st.empty()
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines), language=None)

    # 1. Fetch dữ liệu
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

    # 2. Tạo file Excel → base64
    log("💾 Đang tạo file Excel...")
    fname  = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    buffer = build_excel_buffer(save_mode_val, all_kw_data)
    if not buffer:
        st.error("Không thể tạo file Excel.")
        st.stop()

    excel_b64  = base64.b64encode(buffer.getvalue()).decode('utf-8')
    request_id = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{re.sub(r'[^a-z0-9]', '', user_email.lower())[:10]}"
    log(f"✅ File Excel đã tạo: {fname} ({len(buffer.getvalue())//1024} KB)")

    # 3. Đăng ký request lên Google Apps Script (lưu tạm file + metadata)
    log("☁️  Đang gửi dữ liệu lên Google Apps Script...")
    gas_ok = register_request_to_gas(
        request_id, user_email, keywords,
        results_summary, excel_b64, fname, save_mode_label
    )
    if gas_ok:
        log("✅ Google Apps Script đã nhận và lưu dữ liệu.")
    else:
        log("⚠️  GAS không phản hồi — vẫn tiếp tục gửi Telegram.")

    # 4. Gửi Telegram với nút duyệt
    log("📨 Đang gửi thông báo Telegram...")
    tg_ok, tg_resp = send_telegram_approval(
        user_email, keywords, results_summary,
        excel_b64, fname, save_mode_label, request_id
    )
    if tg_ok:
        log("✅ Đã gửi Telegram thành công!")
    else:
        log(f"❌ Telegram lỗi: {tg_resp}")

    # 5. Thông báo cho user
    st.success(
        f"✅ Yêu cầu đã được gửi! Bạn sẽ nhận kết quả tại **{user_email}** "
        f"sau khi được phê duyệt. Mã yêu cầu: `{request_id}`"
    )
