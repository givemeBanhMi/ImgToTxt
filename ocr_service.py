import openai
import base64
import re
import time
import os
from config_manager import load_config

def encode_image_to_base64(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def extract_passport_data(image_path, api_key, api_base_url, model_name, use_offline=False):
    """
    Sends the image to 9router/OpenAI-compatible API and extracts full passport fields.
    Returns a dictionary with all fields.
    """
    # Offline mode bypasses API key/model validation
    if use_offline:
        try:
            from PIL import Image
            ocr = _get_easyocr()
        except Exception as e:
            return _empty_result(
                "OCR Offline Error: Thiếu thư viện. "
                "Hãy cài đặt: pip install easyocr passporteye pillow"
            )

        _configure_tesseract()

        passporteye_result = None
        try:
            passporteye_result = _extract_with_passporteye(image_path)
        except Exception:
            passporteye_result = None

        try:
            result = ocr.readtext(image_path)
            raw_text_lines = []
            if result:
                for line in result:
                    if len(line) >= 2:
                        text_val = line[1]
                        raw_text_lines.append(text_val)
            raw_text = "\n".join(raw_text_lines)
        except Exception as e:
            return _empty_result(f"OCR Offline Error (EasyOCR): {e}")

        fields = _parse_offline_passport_text(raw_text)
        fields = _merge_offline_results(passporteye_result, fields)
        if not fields["so_passport"] or fields["so_passport"] == "Không tìm thấy":
            return _empty_result("OCR Thất bại (Không tìm thấy số Passport)")
        return fields
    if not api_key:
        return _empty_result("Thiếu API Key")
    if not model_name:
        return _empty_result("Thiếu Model Name")

    try:
        client_kwargs = {"api_key": api_key, "timeout": 45.0}
        if api_base_url:
            client_kwargs["base_url"] = api_base_url

        client = openai.OpenAI(**client_kwargs)
        base64_image = encode_image_to_base64(image_path)

        prompt = (
            "Hãy đọc toàn bộ nội dung trong ảnh hộ chiếu này và trích xuất thông tin. "
            "Trả về KẾT QUẢ CHÍNH XÁC theo đúng định dạng sau (mỗi trường một dòng). "
            "Nếu không tìm thấy thông tin, ghi 'Không tìm thấy':\n\n"
            "Số Passport: [giá trị]\n"
            "Họ tên: [giá trị]\n"
            "Ngày sinh: [giá trị]\n"
            "Giới tính: [giá trị]\n"
            "Quốc tịch: [giá trị]\n"
            "Nơi sinh: [giá trị]\n"
            "Ngày cấp: [giá trị]\n"
            "Ngày hết hạn: [giá trị]\n"
            "Cơ quan cấp: [giá trị]\n"
            "Dòng MRZ: [giá trị]\n\n"
            "Chỉ trả về nội dung theo format trên, không giải thích thêm."
        )

        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=800
        )

        text = response.choices[0].message.content

        # Parse all fields
        fields = {
            "so_passport":   _parse_field(text, "Số Passport"),
            "ho_ten":        _parse_field(text, "Họ tên"),
            "ngay_sinh":     normalize_date(_parse_field(text, "Ngày sinh")),
            "gioi_tinh":     _parse_field(text, "Giới tính"),
            "quoc_tich":     _parse_field(text, "Quốc tịch"),
            "noi_sinh":      _parse_field(text, "Nơi sinh"),
            "ngay_cap":      normalize_date(_parse_field(text, "Ngày cấp")),
            "ngay_het_han":  normalize_date(_parse_field(text, "Ngày hết hạn")),
            "co_quan_cap":   _parse_field(text, "Cơ quan cấp"),
            "dong_mrz":      _parse_field(text, "Dòng MRZ"),
            "trang_thai":    "Thành công"
        }

        if not fields["so_passport"] or fields["so_passport"] == "Không tìm thấy":
            return _empty_result("OCR Thất bại (Không tìm thấy số Passport)")

        return fields

    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg:
            return _empty_result("Lỗi mạng/Quá tải (Rate Limit)")
        return _empty_result(f"Lỗi OCR: {error_msg}")


def normalize_date(date_str):
    import re
    if not date_str:
        return ""
    
    date_str = date_str.strip().upper()
    
    # Replace common month abbreviations
    months = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
        "THG 1": "01", "THG 2": "02", "THG 3": "03", "THG 4": "04",
        "THG 5": "05", "THG 6": "06", "THG 7": "07", "THG 8": "08",
        "THG 9": "09", "THG 10": "10", "THG 11": "11", "THG 12": "12",
    }
    
    # Handle "DD MMM YYYY" format like "14 AUG 1983"
    m_alpha = re.search(r"(\d{1,2})[ \-\/]?([A-Z]{3}|THG \d{1,2})[ \-\/]?(\d{4})", date_str)
    if m_alpha:
        d = m_alpha.group(1).zfill(2)
        m_m = months.get(m_alpha.group(2), "01")
        y = m_alpha.group(3)
        return f"{d}/{m_m}/{y}"
        
    # Handle "YYYYMMDD" format
    m_ymd = re.match(r"^(\d{4})(\d{2})(\d{2})$", date_str)
    if m_ymd:
        return f"{m_ymd.group(3)}/{m_ymd.group(2)}/{m_ymd.group(1)}"
        
    # Handle "YYYY/MM/DD" or "YYYY-MM-DD"
    m_ymd2 = re.match(r"^(\d{4})[-/.](\d{2})[-/.](\d{2})$", date_str)
    if m_ymd2:
        return f"{m_ymd2.group(3)}/{m_ymd2.group(2)}/{m_ymd2.group(1)}"

    # Handle "DD/MM/YYYY" or "DD-MM-YYYY" (standardize separator)
    m_dmy = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", date_str)
    if m_dmy:
        return f"{m_dmy.group(1).zfill(2)}/{m_dmy.group(2).zfill(2)}/{m_dmy.group(3)}"
        
    return date_str

def _parse_field(text, field_name):
    """Extract a single field value from the formatted text."""
    pattern = rf"{re.escape(field_name)}:\s*(.+)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        val = match.group(1).strip()
        if val.lower() in ("không tìm thấy", "none", "n/a", "", "-"):
            return ""
        return val
    return ""


_easyocr_instance = None

def _get_easyocr():
    global _easyocr_instance
    if _easyocr_instance is None:
        import easyocr
        _easyocr_instance = easyocr.Reader(['en'])
    return _easyocr_instance

def _configure_tesseract():
    try:
        import pytesseract
        import os
        from config_manager import load_config
        cfg = load_config()
        tesseract_path = cfg.get("tesseract_path", "").strip()
        possible_paths = []
        if tesseract_path:
            possible_paths.append(tesseract_path)
        possible_paths.extend([
            os.path.join(os.getcwd(), "Tesseract-OCR", "tesseract.exe"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tesseract-OCR", "tesseract.exe"),
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        ])
        for path in possible_paths:
            if os.path.isfile(path):
                pytesseract.pytesseract.tesseract_cmd = path
                os.environ["PATH"] = os.path.dirname(path) + os.pathsep + os.environ.get("PATH", "")
                return path
    except Exception:
        pass
    return ""


def _extract_with_passporteye(image_path):
    try:
        from passporteye import read_mrz
    except ModuleNotFoundError:
        return None

    mrz = read_mrz(image_path)
    if not mrz:
        return None

    data = mrz.to_dict()
    mrz_text = getattr(mrz, "aux", {}).get("method", "")
    raw_lines = getattr(mrz, "mrz", "") or data.get("raw_text", "") or ""
    if not raw_lines:
        raw_lines = "\n".join(filter(None, [data.get("mrz_type", ""), mrz_text]))

    return {
        "so_passport": data.get("number", "") or data.get("document_number", ""),
        "ho_ten": _format_passporteye_name(data),
        "ngay_sinh": _normalize_passporteye_date(data.get("date_of_birth", "")),
        "gioi_tinh": _normalize_sex(data.get("sex", "")),
        "quoc_tich": data.get("nationality", ""),
        "noi_sinh": "",
        "ngay_cap": "",
        "ngay_het_han": _normalize_passporteye_date(data.get("expiration_date", "")),
        "co_quan_cap": "",
        "dong_mrz": raw_lines,
        "trang_thai": "Thành công"
    }


def _format_passporteye_name(data):
    surname = (data.get("surname") or "").replace("<", " ").strip()
    names = (data.get("names") or "").replace("<", " ").strip()
    full = " ".join(part for part in [surname, names] if part).strip()
    return re.sub(r"\s+", " ", full)


def _normalize_passporteye_date(value):
    value = re.sub(r"[^0-9]", "", str(value or ""))
    if len(value) == 6:
        return _normalize_mrz_date(value)
    if len(value) == 8:
        return normalize_date(value)
    return ""


def _clean_extracted_name(name):
    if not name: return ""
    name = name.strip()
    # Remove MRZ filler garbage misread as K, S, < (e.g. "KK S KKSKKKKSKSK")
    # Requires at least 4 of these characters at the end of the string to be safely removed
    name = re.sub(r"(?i)\s*(?:[KS<]\s*){4,}$", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip()

def _merge_offline_results(passporteye_result, raw_result):
    raw_name = _clean_extracted_name(raw_result.get("ho_ten", ""))
    if raw_name:
        raw_result["ho_ten"] = raw_name

    if not passporteye_result:
        return raw_result

    merged = raw_result.copy()
    for key in ["so_passport", "ngay_sinh", "gioi_tinh", "quoc_tich", "ngay_het_han", "dong_mrz"]:
        if passporteye_result.get(key):
            merged[key] = passporteye_result[key]

    passporteye_name = _clean_extracted_name(passporteye_result.get("ho_ten", ""))
    
    if passporteye_name:
        merged["ho_ten"] = passporteye_name
    else:
        merged["ho_ten"] = raw_result.get("ho_ten", "")

    # Fallback: if issue date is still empty but we have an expiry date, infer it again
    if not merged.get("ngay_cap") and merged.get("ngay_het_han"):
        merged["ngay_cap"] = _infer_issue_date_from_expiry(merged["ngay_het_han"])

    return merged


def _parse_offline_passport_text(raw_text):
    """Parse raw Tesseract text from a passport image."""
    text = raw_text or ""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    mrz_lines = _find_mrz_lines(lines)
    mrz1 = mrz_lines[0] if len(mrz_lines) > 0 else ""
    mrz2 = mrz_lines[1] if len(mrz_lines) > 1 else ""

    passport_no = ""
    nationality = ""
    birth_date = ""
    sex = ""
    expiry_date = ""
    full_name = ""

    if mrz2:
        passport_no = re.sub(r"[^A-Z0-9]", "", mrz2[0:9]).replace("O", "0")
        nationality = mrz2[10:13].replace("<", "").strip()
        birth_date = _normalize_mrz_date(mrz2[13:19])
        sex = _normalize_sex(mrz2[20:21])
        expiry_date = _normalize_mrz_date(mrz2[21:27])

    if mrz1:
        name_part = mrz1[5:] if len(mrz1) > 5 else mrz1
        name_part = name_part.replace("0", "O")
        chunks = name_part.split("<<", 1)
        surname = re.sub(r"[^A-Z ]", "", chunks[0].replace("<", " ")).strip()
        given_source = chunks[1].split("<<<", 1)[0] if len(chunks) > 1 else ""
        given = re.sub(r"[^A-Z ]", "", given_source.replace("<", " ")).strip()
        full_name = " ".join(part for part in [surname, given] if part).strip()

    if not passport_no:
        passport_no = _guess_passport_number(text)
    if not full_name:
        full_name = _guess_name(lines)
    if not birth_date:
        birth_date = normalize_date(_guess_birth_date(text))
    if not expiry_date:
        expiry_date = normalize_date(_guess_expiry_date(text))

    issue_date = _guess_issue_date(text)
    if not issue_date:
        issue_date = _infer_issue_date_from_expiry(expiry_date)

    return {
        "so_passport": passport_no,
        "ho_ten": full_name,
        "ngay_sinh": birth_date,
        "gioi_tinh": sex,
        "quoc_tich": nationality,
        "noi_sinh": _clean_place(_guess_place_of_birth(lines)),
        "ngay_cap": issue_date,
        "ngay_het_han": expiry_date,
        "co_quan_cap": _guess_authority(lines),
        "dong_mrz": "\n".join(mrz_lines),
        "trang_thai": "Thành công"
    }


def _find_mrz_lines(lines):
    candidates = []
    for line in lines:
        cleaned = re.sub(r"[^A-Z0-9<]", "", line.upper())
        looks_like_mrz_name = len(cleaned) >= 30 and ("<<" in cleaned or cleaned.startswith(("P<", "PP")))
        looks_like_mrz_data = len(cleaned) >= 30 and re.search(r"[A-Z]{1,2}\d{6,8}\d[A-Z]{3}\d{7}[MF<]\d{6}", cleaned)
        if looks_like_mrz_name or looks_like_mrz_data:
            candidates.append(cleaned)
    return candidates[-2:]


def _normalize_mrz_date(value):
    value = re.sub(r"[^0-9]", "", value or "")
    if len(value) != 6:
        return ""
    yy = int(value[:2])
    year = 1900 + yy if yy >= 50 else 2000 + yy
    return f"{value[4:6]}/{value[2:4]}/{year}"


def _normalize_sex(value):
    value = (value or "").upper()
    if value == "M":
        return "Nam"
    if value == "F":
        return "Nữ"
    return value


def _guess_passport_number(text):
    matches = re.findall(r"\b[A-Z]{1,2}\d{6,8}\b", text.upper())
    return matches[0] if matches else ""


def _guess_name(lines):
    for line in lines:
        cleaned = line.strip()
        if re.match(r"^[A-Z][A-Z\s,.'-]{3,}$", cleaned) and "," in cleaned:
            return cleaned.replace(",", " ").strip()
    return ""


def _guess_birth_date(text):
    match = re.search(r"(\d{1,2}\s+[A-Z]{3}\s+\d{4})", text.upper())
    return match.group(1) if match else ""


def _guess_expiry_date(text):
    matches = re.findall(r"(\d{1,2}\s+[A-Z]{3}\s+\d{4})", text.upper())
    return matches[-1] if matches else ""


def _guess_issue_date(text):
    upper = text.upper()
    dates = [normalize_date(match) for match in re.findall(r"\d{1,2}\s+[A-Z]{3}\s+\d{4}", upper)]
    return dates[-2] if len(dates) >= 2 else ""


def _infer_issue_date_from_expiry(expiry_date):
    try:
        from datetime import datetime, timedelta
        expiry = datetime.strptime(expiry_date, "%d/%m/%Y")
        issue = expiry.replace(year=expiry.year - 10) + timedelta(days=1)
        return issue.strftime("%d/%m/%Y")
    except Exception:
        return ""


def _guess_place_of_birth(lines):
    for line in lines:
        upper = line.upper()
        if "/" in upper and any(word in upper for word in ["ZHE", "JIANG", "BEIJING", "SHANGHAI"]):
            return upper
    return ""


def _clean_place(value):
    upper = (value or "").upper()
    if "ZHE" in upper and ("JIANG" in upper or "JTANG" in upper or "TANG" in upper):
        return "ZHEJIANG"
    if "BEIJING" in upper:
        return "BEIJING"
    if "SHANGHAI" in upper:
        return "SHANGHAI"
    return re.sub(r"[^A-Z ]", "", upper).strip()


def _guess_authority(lines):
    for line in lines:
        upper = line.upper()
        if any(kw in upper for kw in ["IMMIGRATION", "FOREIGN AFFAIRS", "MINISTRY"]):
            return line.strip()
    return ""


def _empty_result(status_msg):
    """Return a dict with all empty fields and an error status."""
    return {
        "so_passport":  "",
        "ho_ten":       "",
        "ngay_sinh":    "",
        "gioi_tinh":    "",
        "quoc_tich":    "",
        "noi_sinh":     "",
        "ngay_cap":     "",
        "ngay_het_han": "",
        "co_quan_cap":  "",
        "dong_mrz":     "",
        "trang_thai":   status_msg
    }
