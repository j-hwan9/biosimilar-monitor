import os

EMAIL_CONFIG = {
    "sender":      os.environ.get("GMAIL_SENDER", "your_gmail@gmail.com"),
    "app_password":os.environ.get("GMAIL_APP_PASSWORD", "xxxx xxxx xxxx xxxx"),
    "recipients":  os.environ.get("GMAIL_RECIPIENTS", "your_email@company.com").split(","),
}
"""
CMS ASP Biosimilar Monitor - FINAL VERSION
=========================================
Samsung Bioepis Biosimilar Market Report 검증 파일

[포함 Molecule]
Medical Benefit 기존 4개 (SB 제품 있음):
  Infliximab, Ranibizumab, Trastuzumab, Denosumab

추가 6개 (SB 제품 없음, 시장 검증용):
  Bevacizumab, Rituximab, Filgrastim, Pegfilgrastim,
  Epoetin alfa, Tocilizumab

[핵심 로직]
- HCPCS 코드 하드코딩 금지 → Short Description 키워드 매핑
- ASP 역산: Originator = PL/1.06, Biosimilar = PL-(RefASP×addon%)
- IRA +8% qualifying: Biosimilar ASP ≤ Reference ASP 조건 검증 후 적용
- Notes 컬럼 보조 활용 (CMS 명시 시)
- 최대 12개 분기 자동 수집

[임상 단위 기준]
  Infliximab    J1745 per 10mg    → 100mg/vial   (×10)
  Ranibizumab   J2778 per 0.1mg   → 0.5mg/inj    (×5)
  Trastuzumab   J9355 per 10mg    → 420mg/vial   (×42)
  Denosumab     J0897 per 1mg     → 120mg/Xgeva  (×120)
  Bevacizumab   J9035 per 10mg    → 400mg/vial   (×40)
  Rituximab     J9312 per 10mg    → 500mg/vial   (×50)
  Filgrastim    J1442 per 1mcg    → 300mcg/vial  (×300)
  Pegfilgrastim J2505 per 0.5mg   → 6mg/syringe  (×12)
  Epoetin alfa  J0885 per 1000u   → 1000u/vial   (×1)
  Tocilizumab   J3262 per 1mg     → 400mg/vial   (×400)
"""
import io
import zipfile
import smtplib
import requests
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime

# ============================================================
# 설정 영역
# ============================================================
EMAIL_CONFIG = {
    "sender":      "jhwang1637@gmail.com",
    "app_password":"ifpv chtu wkjt ebtt",
    "recipients":  ["jhwang1637@gmail.com"],
}

# 최신 12개 분기 (과거→현재)
QUARTERS = [
    {"label": "2023 Q3 (Jul)", "url": "https://www.cms.gov/files/zip/july-2023-asp-pricing-file.zip"},
    {"label": "2023 Q4 (Oct)", "url": "https://www.cms.gov/files/zip/october-2023-asp-pricing-file.zip"},
    {"label": "2024 Q1 (Jan)", "url": "https://www.cms.gov/files/zip/january-2024-asp-pricing-file.zip"},
    {"label": "2024 Q2 (Apr)", "url": "https://www.cms.gov/files/zip/april-2024-asp-pricing-file.zip"},
    {"label": "2024 Q3 (Jul)", "url": "https://www.cms.gov/files/zip/july-2024-asp-pricing-file.zip"},
    {"label": "2024 Q4 (Oct)", "url": "https://www.cms.gov/files/zip/october-2024-asp-pricing-file.zip"},
    {"label": "2025 Q1 (Jan)", "url": "https://www.cms.gov/files/zip/january-2025-asp-pricing-file-03/11/25-final-file.zip"},
    {"label": "2025 Q2 (Apr)", "url": "https://www.cms.gov/files/zip/april-2025-asp-pricing-file.zip"},
    {"label": "2025 Q3 (Jul)", "url": "https://www.cms.gov/files/zip/july-2025-asp-pricing-file.zip"},
    {"label": "2025 Q4 (Oct)", "url": "https://www.cms.gov/files/zip/october-2025-asp-pricing-final-file.zip"},
    {"label": "2026 Q1 (Jan)", "url": "https://www.cms.gov/files/zip/january-2026-medicare-part-b-payment-limit-files.zip"},
    {"label": "2026 Q2 (Apr)", "url": "https://www.cms.gov/files/zip/april-2026-medicare-part-b-payment-limit-files-03-19-2026-final-file.zip"},
]

CROSSWALK_URL = "https://www.cms.gov/files/zip/april-2026-ndc-hcpcs-crosswalk-03-19-2026-final-file.zip"

# ============================================================
# 제품 DB
# hcpcs_fixed   : Originator J코드 (불변)
# desc_keywords : Short Description 키워드 (소문자)
# desc_exclude  : 이 키워드 포함 시 제외
# display_mult  : CMS 1unit → 임상 기준 환산 배수
# has_sb        : Samsung Bioepis 제품 여부
# ============================================================
MOLECULES = {
    # ── Samsung Bioepis 제품 있는 molecule ──
    "Infliximab": {
        "benefit": "Medical Benefit — IV Infusion",
        "display_dose": "100mg/vial", "display_mult": 10, "has_sb": True,
        "originator": {"hcpcs_fixed": "J1745", "brand": "Remicade",
                       "company": "J&J (Janssen)", "fda_name": "REMICADE",
                       "desc_keywords": ["infliximab", "remicade"], "desc_exclude": ["biosim"]},
        "biosimilars": [
            {"brand": "Renflexis",  "suffix": "infliximab-abda",    "company": "Samsung Bioepis", "is_sb": True,
             "desc_keywords": ["renflexis"],  "desc_exclude": [], "fda_name": "RENFLEXIS"},
            {"brand": "Inflectra",  "suffix": "infliximab-dyyb",    "company": "Pfizer/Celltrion","is_sb": False,
             "desc_keywords": ["inflectra"],  "desc_exclude": [], "fda_name": "INFLECTRA"},
            {"brand": "Avsola",     "suffix": "infliximab-axxq",    "company": "Amgen",           "is_sb": False,
             "desc_keywords": ["avsola"],     "desc_exclude": [], "fda_name": "AVSOLA"},
            {"brand": "Ixifi",      "suffix": "infliximab-qbtx",    "company": "Pfizer",          "is_sb": False,
             "desc_keywords": ["ixifi"],      "desc_exclude": [], "fda_name": "IXIFI"},
            {"brand": "Zymfentra",  "suffix": "infliximab-dyyb SC", "company": "Celltrion",       "is_sb": False,
             "desc_keywords": ["zymfentra"],  "desc_exclude": [], "fda_name": "ZYMFENTRA"},
        ],
    },
    "Ranibizumab": {
        "benefit": "Medical Benefit — Intravitreal Injection",
        "display_dose": "0.5mg/injection", "display_mult": 5, "has_sb": True,
        "originator": {"hcpcs_fixed": "J2778", "brand": "Lucentis",
                       "company": "Genentech (Roche)", "fda_name": "LUCENTIS",
                       "desc_keywords": ["lucentis", "ranibizumab"], "desc_exclude": ["biosim","nuna","eqrn","susvimo"]},
        "biosimilars": [
            {"brand": "Byooviz", "suffix": "ranibizumab-nuna", "company": "Samsung Bioepis", "is_sb": True,
             "desc_keywords": ["byooviz","nuna"], "desc_exclude": [], "fda_name": "BYOOVIZ"},
            {"brand": "Cimerli", "suffix": "ranibizumab-eqrn", "company": "Coherus",         "is_sb": False,
             "desc_keywords": ["cimerli","eqrn"], "desc_exclude": [], "fda_name": "CIMERLI"},
        ],
    },
    "Trastuzumab": {
        "benefit": "Medical Benefit — IV Infusion",
        "display_dose": "420mg/vial", "display_mult": 42, "has_sb": True,
        "originator": {"hcpcs_fixed": "J9355", "brand": "Herceptin",
                       "company": "Genentech (Roche)", "fda_name": "HERCEPTIN",
                       "desc_keywords": ["herceptin","trastuzumab"], "desc_exclude": ["biosim","dttb","qyyp","pkrb","anns","dkst","strf"]},
        "biosimilars": [
            {"brand": "Ontruzant", "suffix": "trastuzumab-dttb", "company": "Samsung Bioepis", "is_sb": True,
             "desc_keywords": ["ontruzant","dttb"], "desc_exclude": [], "fda_name": "ONTRUZANT"},
            {"brand": "Herzuma",   "suffix": "trastuzumab-pkrb", "company": "Celltrion/Teva",  "is_sb": False,
             "desc_keywords": ["herzuma","pkrb"],   "desc_exclude": [], "fda_name": "HERZUMA"},
            {"brand": "Ogivri",    "suffix": "trastuzumab-dkst", "company": "Viatris",         "is_sb": False,
             "desc_keywords": ["ogivri","dkst"],    "desc_exclude": [], "fda_name": "OGIVRI"},
            {"brand": "Trazimera", "suffix": "trastuzumab-qyyp", "company": "Pfizer",          "is_sb": False,
             "desc_keywords": ["trazimera","qyyp"], "desc_exclude": [], "fda_name": "TRAZIMERA"},
            {"brand": "Kanjinti",  "suffix": "trastuzumab-anns", "company": "Amgen",           "is_sb": False,
             "desc_keywords": ["kanjinti","anns"],  "desc_exclude": [], "fda_name": "KANJINTI"},
            {"brand": "Hercessi",  "suffix": "trastuzumab-strf", "company": "Accord/Intas",    "is_sb": False,
             "desc_keywords": ["hercessi","strf"],  "desc_exclude": [], "fda_name": "HERCESSI"},
        ],
    },
    "Denosumab": {
        "benefit": "Medical Benefit — SC Injection (Office)",
        "display_dose": "120mg/vial (Xgeva)", "display_mult": 120, "has_sb": True,
        "originator": {"hcpcs_fixed": "J0897", "brand": "Prolia / Xgeva",
                       "company": "Amgen", "fda_name": "PROLIA",
                       "desc_keywords": ["prolia","xgeva","denosumab"], "desc_exclude": ["biosim","dssb","bbdz","bmwo","bnht","kyqq","nxxp"]},
        "biosimilars": [
            {"brand": "Ospomyv/Xbryk",     "suffix": "denosumab-dssb", "company": "Samsung Bioepis", "is_sb": True,
             "desc_keywords": ["ospomyv","xbryk","dssb"],  "desc_exclude": [], "fda_name": "OSPOMYV"},
            {"brand": "Jubbonti/Wyost",    "suffix": "denosumab-bbdz", "company": "Sandoz",          "is_sb": False,
             "desc_keywords": ["jubbonti","wyost","bbdz"], "desc_exclude": [], "fda_name": "JUBBONTI"},
            {"brand": "Stoboclo/Osenvelt", "suffix": "denosumab-bmwo", "company": "Celltrion",       "is_sb": False,
             "desc_keywords": ["stoboclo","osenvelt","bmwo"], "desc_exclude": [], "fda_name": "STOBOCLO"},
        ],
    },
    # ── Samsung Bioepis 제품 없는 molecule (시장 검증용) ──
    "Bevacizumab": {
        "benefit": "Medical Benefit — IV Infusion",
        "display_dose": "400mg/vial", "display_mult": 40, "has_sb": False,
        "originator": {"hcpcs_fixed": "J9035", "brand": "Avastin",
                       "company": "Genentech (Roche)", "fda_name": "AVASTIN",
                       "desc_keywords": ["bevacizumab","avastin"], "desc_exclude": ["biosim","awwb","bvzr","maly","adcd","tnjn","nwgd"]},
        "biosimilars": [
            {"brand": "Mvasi",    "suffix": "bevacizumab-awwb", "company": "Amgen",      "is_sb": False,
             "desc_keywords": ["mvasi","awwb"],   "desc_exclude": [], "fda_name": "MVASI"},
            {"brand": "Zirabev",  "suffix": "bevacizumab-bvzr", "company": "Pfizer",     "is_sb": False,
             "desc_keywords": ["zirabev","bvzr"], "desc_exclude": [], "fda_name": "ZIRABEV"},
            {"brand": "Alymsys",  "suffix": "bevacizumab-maly", "company": "Amneal",     "is_sb": False,
             "desc_keywords": ["alymsys","maly"], "desc_exclude": [], "fda_name": "ALYMSYS"},
            {"brand": "Vegzelma", "suffix": "bevacizumab-adcd", "company": "Celltrion",  "is_sb": False,
             "desc_keywords": ["vegzelma","adcd"],"desc_exclude": [], "fda_name": "VEGZELMA"},
            {"brand": "Avzivi",   "suffix": "bevacizumab-tnjn", "company": "Sandoz/Bio-Thera", "is_sb": False,
             "desc_keywords": ["avzivi","tnjn"],  "desc_exclude": [], "fda_name": "AVZIVI"},
            {"brand": "Jobevne",  "suffix": "bevacizumab-nwgd", "company": "Biocon",     "is_sb": False,
             "desc_keywords": ["jobevne","nwgd"], "desc_exclude": [], "fda_name": "JOBEVNE"},
        ],
    },
    "Rituximab": {
        "benefit": "Medical Benefit — IV Infusion",
        "display_dose": "500mg/vial", "display_mult": 50, "has_sb": False,
        "originator": {"hcpcs_fixed": "J9312", "brand": "Rituxan",
                       "company": "Genentech/Biogen", "fda_name": "RITUXAN",
                       "desc_keywords": ["rituximab","rituxan"], "desc_exclude": ["biosim","abbs","pvvr","arrx"]},
        "biosimilars": [
            {"brand": "Truxima", "suffix": "rituximab-abbs", "company": "Celltrion/Teva", "is_sb": False,
             "desc_keywords": ["truxima","abbs"], "desc_exclude": [], "fda_name": "TRUXIMA"},
            {"brand": "Ruxience","suffix": "rituximab-pvvr", "company": "Pfizer",         "is_sb": False,
             "desc_keywords": ["ruxience","pvvr"],"desc_exclude": [], "fda_name": "RUXIENCE"},
            {"brand": "Riabni",  "suffix": "rituximab-arrx", "company": "Amgen",          "is_sb": False,
             "desc_keywords": ["riabni","arrx"],  "desc_exclude": [], "fda_name": "RIABNI"},
        ],
    },
    "Filgrastim": {
        "benefit": "Medical Benefit — SC/IV Injection",
        "display_dose": "300mcg/vial", "display_mult": 300, "has_sb": False,
        "originator": {"hcpcs_fixed": "J1442", "brand": "Neupogen",
                       "company": "Amgen", "fda_name": "NEUPOGEN",
                       "desc_keywords": ["filgrastim","neupogen"], "desc_exclude": ["biosim","sndz","aafi","ayow","txid"]},
        "biosimilars": [
            {"brand": "Zarxio",  "suffix": "filgrastim-sndz", "company": "Sandoz",         "is_sb": False,
             "desc_keywords": ["zarxio","sndz"],  "desc_exclude": [], "fda_name": "ZARXIO"},
            {"brand": "Nivestym","suffix": "filgrastim-aafi", "company": "Pfizer/Hospira",  "is_sb": False,
             "desc_keywords": ["nivestym","aafi"],"desc_exclude": [], "fda_name": "NIVESTYM"},
            {"brand": "Releuko", "suffix": "filgrastim-ayow", "company": "Amneal/Kashiv",   "is_sb": False,
             "desc_keywords": ["releuko","ayow"], "desc_exclude": [], "fda_name": "RELEUKO"},
            {"brand": "Nypozi",  "suffix": "filgrastim-txid", "company": "Tanvex",          "is_sb": False,
             "desc_keywords": ["nypozi","txid"],  "desc_exclude": [], "fda_name": "NYPOZI"},
        ],
    },
    "Pegfilgrastim": {
    "benefit": "Medical Benefit — SC Injection",
    "display_dose": "6mg/syringe", "display_mult": 12, "has_sb": False,
    "originator": {"hcpcs_fixed": "J2506", "brand": "Neulasta",
                   "company": "Amgen", "fda_name": "NEULASTA",
                   "desc_keywords": ["pegfilgrastim","neulasta"], "desc_exclude": ["biosim","jmdb","cbqv","bmez","apgf","fpgk","pbbk","unne"]},
        "biosimilars": [
            {"brand": "Fulphila",   "suffix": "pegfilgrastim-jmdb", "company": "Biocon",          "is_sb": False,
             "desc_keywords": ["fulphila","jmdb"],   "desc_exclude": [], "fda_name": "FULPHILA"},
            {"brand": "Udenyca",    "suffix": "pegfilgrastim-cbqv", "company": "Coherus",          "is_sb": False,
             "desc_keywords": ["udenyca","cbqv"],    "desc_exclude": [], "fda_name": "UDENYCA"},
            {"brand": "Ziextenzo",  "suffix": "pegfilgrastim-bmez", "company": "Sandoz",           "is_sb": False,
             "desc_keywords": ["ziextenzo","bmez"],  "desc_exclude": [], "fda_name": "ZIEXTENZO"},
            {"brand": "Nyvepria",   "suffix": "pegfilgrastim-apgf", "company": "Pfizer/Hospira",   "is_sb": False,
             "desc_keywords": ["nyvepria","apgf"],   "desc_exclude": [], "fda_name": "NYVEPRIA"},
            {"brand": "Stimufend",  "suffix": "pegfilgrastim-fpgk", "company": "Fresenius Kabi",   "is_sb": False,
             "desc_keywords": ["stimufend","fpgk"],  "desc_exclude": [], "fda_name": "STIMUFEND"},
            {"brand": "Fylnetra",   "suffix": "pegfilgrastim-pbbk", "company": "Amneal/Kashiv",    "is_sb": False,
             "desc_keywords": ["fylnetra","pbbk"],   "desc_exclude": [], "fda_name": "FYLNETRA"},
        ],
    },
    "Epoetin alfa": {
        "benefit": "Medical Benefit — SC/IV Injection (non-ESRD)",
        "display_dose": "40,000 units/vial", "display_mult": 40, "has_sb": False,
        "originator": {"hcpcs_fixed": "J0885", "brand": "Epogen/Procrit",
                       "company": "Amgen/J&J", "fda_name": "EPOGEN",
                       "desc_keywords": ["epoetin alfa","epogen","procrit"], "desc_exclude": ["biosim","epbx","esrd"]},
        "biosimilars": [
            {"brand": "Retacrit", "suffix": "epoetin alfa-epbx", "company": "Pfizer/Hospira", "is_sb": False,
 "desc_keywords": ["retacrit","epbx"], "desc_exclude": ["esrd on dialysis"], "fda_name": "RETACRIT"},
        ],
    },
    "Tocilizumab": {
        "benefit": "Medical Benefit — IV Infusion / SC Injection",
        "display_dose": "400mg/vial", "display_mult": 400, "has_sb": False,
        "originator": {"hcpcs_fixed": "J3262", "brand": "Actemra",
                       "company": "Genentech (Roche)", "fda_name": "ACTEMRA",
                       "desc_keywords": ["tocilizumab","actemra"], "desc_exclude": ["biosim","bavi","aazg","anoh"]},
        "biosimilars": [
            {"brand": "Tofidence", "suffix": "tocilizumab-bavi", "company": "Bio-Thera/Organon",  "is_sb": False,
             "desc_keywords": ["tofidence","bavi"], "desc_exclude": [], "fda_name": "TOFIDENCE"},
            {"brand": "Tyenne",    "suffix": "tocilizumab-aazg", "company": "Fresenius Kabi",      "is_sb": False,
             "desc_keywords": ["tyenne","aazg"],    "desc_exclude": [], "fda_name": "TYENNE"},
            {"brand": "Avtozma",   "suffix": "tocilizumab-anoh", "company": "Celltrion",           "is_sb": False,
             "desc_keywords": ["avtozma","anoh"],   "desc_exclude": [], "fda_name": "AVTOZMA"},
        ],
    },
}

EXCLUDED_MOLECULES = [
    {"molecule": "Adalimumab",  "brand": "Humira",  "reason": "Pharmacy Benefit 주도 (SC 자가투여) — CMS Part B ASP 해당 없음"},
    {"molecule": "Ustekinumab", "brand": "Stelara", "reason": "Pharmacy Benefit 주도 (SC 자가투여) — CMS Part B ASP 해당 없음"},
    {"molecule": "Etanercept",  "brand": "Enbrel",  "reason": "특허 보호로 미국 출시 불가 (2029년까지)"},
    {"molecule": "Aflibercept", "brand": "Eylea",   "reason": "Samsung Bioepis Opuviz 소송으로 미국 출시 금지 (2027년 예정)"},
]


# ============================================================
# LAYER 1-A: FDA 마케팅 상태 검증
# ============================================================
def check_fda_marketing_status(fda_name: str) -> str:
    try:
        url = (f"https://api.fda.gov/drug/drugsfda.json"
               f"?search=products.brand_name:\"{fda_name}\"&limit=5")
        resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code != 200:
            return "unknown"
        for r in resp.json().get("results", []):
            for prod in r.get("products", []):
                if fda_name.upper() in prod.get("brand_name", "").upper():
                    st = prod.get("marketing_status", "").lower()
                    if "prescription" in st: return "active"
                    if "discontinued" in st: return "discontinued"
        return "unknown"
    except Exception:
        return "unknown"


def validate_products() -> dict:
    print("\n[LAYER 1-A] FDA 마케팅 상태 검증 중...")
    validation = {}
    for mol_name, mol_info in MOLECULES.items():
        orig = mol_info["originator"]
        st = check_fda_marketing_status(orig["fda_name"])
        validation[orig["brand"]] = {"status": st}
        for bs in mol_info["biosimilars"]:
            st = check_fda_marketing_status(bs["fda_name"])
            validation[bs["brand"]] = {"status": st}
    print(f"  -> {len(validation)}개 제품 검증 완료")
    return validation


# ============================================================
# LAYER 1-B: CMS Crosswalk 신규 코드 감지
# ============================================================
def detect_new_hcpcs_codes() -> list:
    print("\n[LAYER 1-B] CMS Crosswalk 신규 코드 스캔 중...")
    new_codes = []
    target_inns = {
        "infliximab","ranibizumab","trastuzumab","denosumab",
        "bevacizumab","rituximab","filgrastim","pegfilgrastim",
        "epoetin","tocilizumab","pembrolizumab"
    }
    known = set()
    for mol in MOLECULES.values():
        known.add(mol["originator"]["hcpcs_fixed"])
    try:
        resp = requests.get(CROSSWALK_URL, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        z = zipfile.ZipFile(io.BytesIO(resp.content))
        candidates = [f for f in z.namelist()
                      if (f.endswith(".xls") or f.endswith(".xlsx")) and "508" not in f.lower()]
        if not candidates: return []
        data = z.read(candidates[0])
        engine = "xlrd" if candidates[0].endswith(".xls") else "openpyxl"
        df = pd.read_excel(io.BytesIO(data), header=None, engine=engine, dtype=str)
        for _, row in df.iterrows():
            row_text = " ".join(str(v).lower() for v in row.values if str(v) != "nan")
            if not any(inn in row_text for inn in target_inns): continue
            for val in row.values:
                code = str(val).strip().upper()
                if len(code) == 5 and code[0] in ("J","Q") and code[1:].isdigit() and code not in known:
                    new_codes.append({"code": code, "context": row_text[:120]})
                    known.add(code)
        if new_codes:
            print(f"  !! 신규 코드 {len(new_codes)}개 발견: {[c['code'] for c in new_codes]}")
        else:
            print("  -> 신규 코드 없음")
    except Exception as e:
        print(f"  !! 오류: {e}")
    return new_codes


# ============================================================
# 2. CMS ASP 파일 파싱 (header=8 고정, 키워드 매핑)
# ============================================================
def download_and_parse(quarter: dict) -> tuple:
    diag = []
    try:
        resp = requests.get(quarter["url"], timeout=90, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
    except Exception as e:
        diag.append({"level": "error", "msg": f"{quarter['label']} 다운로드 실패: {e}"})
        return {}, diag

    try:
        z = zipfile.ZipFile(io.BytesIO(resp.content))
        candidates = [f for f in z.namelist()
                      if (f.endswith(".xls") or f.endswith(".xlsx"))
                      and "payable" not in f.lower() and "508" not in f.lower()
                      and "crosswalk" not in f.lower() and "noc" not in f.lower()]
        if not candidates:
            diag.append({"level": "error", "msg": f"{quarter['label']} Payment Limit 파일 없음"})
            return {}, diag

        target = candidates[0]
        data   = z.read(target)
        engine = "xlrd" if target.endswith(".xls") else "openpyxl"

        # header=8 우선, 실패 시 스캔 방식
        try:
            df = pd.read_excel(io.BytesIO(data), header=8, engine=engine, dtype=str)
            cols = list(df.columns)
            hcpcs_col = next((i for i, c in enumerate(cols) if "HCPCS" in str(c).upper() and "CODE" in str(c).upper()), 0)
            desc_col  = next((i for i, c in enumerate(cols) if "DESC" in str(c).upper()), 1)
            limit_col = next((i for i, c in enumerate(cols) if "PAYMENT" in str(c).upper() and "LIMIT" in str(c).upper()), 3)
            notes_col = next((i for i, c in enumerate(cols) if "NOTE" in str(c).upper()), None)
        except Exception:
            df = pd.read_excel(io.BytesIO(data), header=None, engine=engine, dtype=str)
            hcpcs_col, desc_col, limit_col, notes_col = 0, 1, 3, 10

        # 전체 코드 인덱스 빌드
        raw_index = {}
        for _, row in df.iterrows():
            try:
                code  = str(row.iloc[hcpcs_col]).strip().upper()
                pl    = float(str(row.iloc[limit_col]).replace(",","").strip())
                desc  = str(row.iloc[desc_col]).strip()
                notes = str(row.iloc[notes_col]).strip() if notes_col is not None and notes_col < len(row) else ""
                notes = "" if notes in ("nan","NaN") else notes
                if len(code) == 5 and code[0] in ("J","Q") and code[1:].isdigit() and pl > 0:
                    raw_index[code] = {"pl": pl, "desc": desc, "notes": notes,
                                       "addon_pct_cms": 8 if "8%" in notes else 6}
            except Exception:
                continue

        diag.append({"level": "info", "msg": f"{quarter['label']}: {len(raw_index)}개 코드 로드"})

        # 키워드 매핑으로 제품 찾기
        brand_data = {}
        for mol_name, mol_info in MOLECULES.items():
            orig = mol_info["originator"]
            fixed = orig["hcpcs_fixed"]
            if fixed in raw_index:
                brand_data[orig["brand"]] = {**raw_index[fixed], "hcpcs": fixed}
            else:
                diag.append({"level": "warn",
                             "msg": f"{quarter['label']} — {orig['brand']} ({fixed}) 없음"})

            for bs in mol_info["biosimilars"]:
                kws  = bs["desc_keywords"]
                excl = bs["desc_exclude"]
                for code, entry in raw_index.items():
                    dl = entry["desc"].lower()
                    if any(kw in dl for kw in kws) and not any(ex in dl for ex in excl):
                        brand_data[bs["brand"]] = {**entry, "hcpcs": code}
                        break

        return brand_data, diag

    except Exception as e:
        diag.append({"level": "error", "msg": f"{quarter['label']} 파싱 오류: {e}"})
        return {}, diag


# ============================================================
# 3. ASP 역산 (IRA qualifying 조건 자동 검증)
# ============================================================
def calc_asp(brand: str, mol_info: dict, brand_data: dict) -> dict | None:
    entry = brand_data.get(brand)
    if entry is None: return None

    pl    = entry["pl"]
    notes = entry["notes"]
    hcpcs = entry["hcpcs"]
    orig_brand = mol_info["originator"]["brand"]

    if brand == orig_brand:
        # Originator
        return {"asp": round(pl/1.06, 6), "payment_limit": pl,
                "addon_pct": 6, "notes": notes, "hcpcs": hcpcs,
                "asp_exact": True, "desc": entry["desc"]}
    else:
        # Biosimilar: IRA qualifying 조건 자동 검증
        orig_entry = brand_data.get(orig_brand)
        if orig_entry is None:
            return {"asp": pl, "payment_limit": pl, "addon_pct": 6,
                    "notes": notes, "hcpcs": hcpcs, "asp_exact": False, "desc": entry["desc"]}

        ref_asp = orig_entry["pl"] / 1.06

        # Step 1: +6% 가정으로 Biosimilar ASP 역산
        bs_asp_6pct = pl - (ref_asp * 0.06)

        # Step 2: IRA qualifying 조건 확인 (Biosimilar ASP ≤ Reference ASP)
        if bs_asp_6pct <= ref_asp:
            # qualifying → +8% 적용
            bs_asp = pl - (ref_asp * 0.08)
            addon  = 8
        else:
            # non-qualifying → +6% 유지
            bs_asp = bs_asp_6pct
            addon  = 6

        # CMS Notes에 명시된 경우 검증 (불일치 시 로그)
        cms_addon = entry.get("addon_pct_cms", 6)

        return {"asp": round(bs_asp, 6), "payment_limit": pl,
                "addon_pct": addon, "addon_pct_cms": cms_addon,
                "notes": notes, "hcpcs": hcpcs,
                "asp_exact": True, "desc": entry["desc"],
                "ira_qualifying": bs_asp_6pct <= ref_asp}


# ============================================================
# 4. 데이터 수집
# ============================================================
def collect_asp_data(validation: dict) -> tuple:
    print("\n[2/4] CMS ASP 파일 다운로드 및 파싱 (최대 12개 분기)")
    all_diag = []
    quarter_brand_data = []

    for q in QUARTERS:
        bd, diag = download_and_parse(q)
        quarter_brand_data.append(bd)
        all_diag.extend(diag)
        ok = "✅" if bd else "❌"
        print(f"  {ok} {q['label']}: {len(bd)}개 제품 매핑")

    quarter_labels = [q["label"] for q in QUARTERS]
    result = {}

    for mol_name, mol_info in MOLECULES.items():
        orig = mol_info["originator"]
        mult = mol_info["display_mult"]
        products = []

        orig_results = [calc_asp(orig["brand"], mol_info, bd) for bd in quarter_brand_data]
        products.append({
            "brand": orig["brand"], "company": orig["company"],
            "suffix": orig["hcpcs_fixed"],
            "is_originator": True, "is_sb": False,
            "fda_status": validation.get(orig["brand"], {}).get("status", "unknown"),
            "asp_results": orig_results, "mult": mult,
        })

        for bs in mol_info["biosimilars"]:
            fda_st = validation.get(bs["brand"], {}).get("status", "unknown")
            if fda_st == "discontinued":
                print(f"  [제외] {bs['brand']} — FDA: discontinued")
                continue
            bs_results = [calc_asp(bs["brand"], mol_info, bd) for bd in quarter_brand_data]
            if any(r is not None for r in bs_results):
                products.append({
                    "brand": bs["brand"], "company": bs["company"],
                    "suffix": bs["suffix"],
                    "is_originator": False, "is_sb": bs.get("is_sb", False),
                    "fda_status": fda_st,
                    "asp_results": bs_results, "mult": mult,
                })

        result[mol_name] = {
            "benefit":      mol_info["benefit"],
            "display_dose": mol_info["display_dose"],
            "has_sb":       mol_info["has_sb"],
            "quarters":     quarter_labels,
            "products":     products,
        }

    return result, all_diag


# ============================================================
# 5. 그래프 (과거→현재, 선 끝 라벨)
# ============================================================
def make_chart(mol_name: str, mol_data: dict) -> bytes:
    quarters     = mol_data["quarters"]
    products     = mol_data["products"]
    display_dose = mol_data["display_dose"]
    has_sb       = mol_data["has_sb"]
    n_q          = len(quarters)

    fig, ax = plt.subplots(figsize=(13, 5.5))
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#F8F9FA")

    sb_colors    = ["#1565C0", "#1976D2", "#42A5F5"]
    other_colors = ["#607D8B","#78909C","#90A4AE","#546E7A","#B0BEC5","#455A64","#37474F","#263238"]
    sb_i = other_i = 0
    lines_meta = []

    for prod in products:
        xs, ys = [], []
        for i, r in enumerate(prod["asp_results"]):
            if r is not None:
                xs.append(i)
                ys.append(round(r["asp"] * prod["mult"], 2))
        if not xs: continue

        if prod["is_originator"]:
            color, lw, ls, ms = "#C62828", 3, "-", 8
        elif prod["is_sb"]:
            color = sb_colors[sb_i % len(sb_colors)]; sb_i += 1
            lw, ls, ms = 2.5, "--", 7
        else:
            color = other_colors[other_i % len(other_colors)]; other_i += 1
            lw, ls, ms = 1.5, ":", 5

        ax.plot(xs, ys, marker="o", linewidth=lw, linestyle=ls,
                color=color, markersize=ms, zorder=3)
        lines_meta.append((xs[-1], ys[-1],
                           f"{prod['brand']}\n({prod['company']})",
                           color, prod["is_originator"] or prod["is_sb"]))

    # 라벨 겹침 방지
    lines_meta.sort(key=lambda m: m[1], reverse=True)
    placed = []
    y_vals = [m[1] for m in lines_meta]
    min_gap = (max(y_vals) - min(y_vals)) * 0.09 if len(y_vals) > 1 and max(y_vals) != min(y_vals) else 20

    for (x_end, y_end, txt, color, bold) in lines_meta:
        y_place = y_end
        for py in placed:
            if abs(y_place - py) < min_gap:
                y_place = py + min_gap
        placed.append(y_place)
        ax.annotate(txt, xy=(x_end, y_end), xytext=(x_end + 0.2, y_place),
                    fontsize=7, color=color,
                    fontweight="bold" if bold else "normal", va="center",
                    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=color, alpha=0.85, lw=0.8),
                    arrowprops=dict(arrowstyle="-", color=color, lw=0.7, alpha=0.5)
                    if abs(y_place - y_end) > 2 else None)

    ax.set_xticks(range(n_q))
    ax.set_xticklabels(quarters, fontsize=8, rotation=25, ha="right")
    ax.set_xlim(-0.3, n_q - 1 + 3.5)

    sb_tag = "  ★ Samsung Bioepis 제품 포함" if has_sb else "  (시장 검증용 — SB 제품 없음)"
    ax.set_title(f"{mol_name}  —  ASP  ({display_dose})\n{mol_data['benefit']}{sb_tag}",
                 fontsize=11, fontweight="bold", pad=10)
    ax.set_ylabel(f"USD ($/{display_dose})", fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.35, zorder=0)
    ax.spines[["top","right"]].set_visible(False)
    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ============================================================
# 6. HTML 이메일
# ============================================================
def build_diag_html(all_diag: list) -> str:
    errors = [d for d in all_diag if d["level"] == "error"]
    warns  = [d for d in all_diag if d["level"] == "warn"]
    if not errors and not warns: return ""
    html = '<div style="margin:14px 28px;">'
    if errors:
        items = "".join(f"<li>{d['msg']}</li>" for d in errors)
        html += f'<div style="background:#FFEBEE;border:1px solid #EF9A9A;border-radius:6px;padding:12px 16px;margin-bottom:8px;"><strong>❌ 실행 오류</strong><ul style="margin:4px 0 0;padding-left:18px;font-size:12px;">{items}</ul></div>'
    if warns:
        items = "".join(f"<li>{d['msg']}</li>" for d in warns)
        html += f'<div style="background:#FFF3CD;border:1px solid #FFC107;border-radius:6px;padding:12px 16px;margin-bottom:8px;"><strong>⚠️ 코드 매핑 주의</strong><ul style="margin:4px 0 0;padding-left:18px;font-size:12px;">{items}</ul></div>'
    html += "</div>"
    return html


def build_email_html(asp_data: dict, chart_cids: list,
                     validation: dict, new_codes: list, all_diag: list) -> str:
    today = datetime.now().strftime("%Y년 %m월 %d일 %H:%M")

    fda_disc = [b for b, v in validation.items() if v["status"] == "discontinued"]
    fda_unk  = [b for b, v in validation.items() if v["status"] == "unknown"]
    val_html = ""
    if fda_disc:
        items = "".join(f"<li>{b}</li>" for b in fda_disc)
        val_html += f'<div style="background:#FFF3CD;border:1px solid #FFC107;border-radius:6px;padding:10px 14px;margin-bottom:6px;"><strong>⚠️ FDA Discontinued — 제외됨</strong><ul style="margin:4px 0 0;padding-left:16px;">{items}</ul></div>'
    if fda_unk:
        items = "".join(f"<li>{b}</li>" for b in fda_unk)
        val_html += f'<div style="background:#E8F4FD;border:1px solid #90CAF9;border-radius:6px;padding:10px 14px;margin-bottom:6px;"><strong>ℹ️ FDA 상태 미확인</strong><ul style="margin:4px 0 0;padding-left:16px;">{items}</ul></div>'

    new_code_html = ""
    if new_codes:
        items = "".join(f"<li><strong>{c['code']}</strong> — {c['context']}</li>" for c in new_codes)
        new_code_html = f'<div style="background:#F3E5F5;border:2px solid #7B1FA2;border-radius:8px;padding:12px 16px;margin:10px 0;"><strong style="color:#7B1FA2;">🆕 신규 HCPCS 코드 감지</strong><ul style="margin:6px 0 0;padding-left:16px;font-size:12px;">{items}</ul></div>'

    diag_html = build_diag_html(all_diag)
    excluded_li = "".join(
        f"<li><strong>{e['molecule']} ({e['brand']})</strong>: {e['reason']}</li>"
        for e in EXCLUDED_MOLECULES)

    # 섹션 구분: SB 제품 있는 것 / 시장 검증용
    sb_mols    = [(n, d) for n, d in asp_data.items() if d["has_sb"]]
    nosb_mols  = [(n, d) for n, d in asp_data.items() if not d["has_sb"]]

    html = f"""
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:980px;margin:0 auto;color:#1a1a1a;">
      <div style="background:#003087;color:white;padding:22px 28px;border-radius:10px 10px 0 0;">
        <h2 style="margin:0;font-size:20px;">Biosimilar ASP Monitor — Samsung Bioepis Market Report 검증 파일</h2>
        <p style="margin:6px 0 0;opacity:.8;font-size:13px;">
          {today} &nbsp;|&nbsp; CMS Medicare Part B &nbsp;|&nbsp; 최대 12개 분기 (과거→현재)
        </p>
      </div>
      <div style="background:#EEF2F7;padding:8px 28px 10px;font-size:12px;color:#444;border-bottom:1px solid #D0D7E3;">
        <strong>ASP 역산:</strong>
        Originator = PL÷1.06 &nbsp;|&nbsp;
        Biosimilar = PL−(RefASP×addon%) &nbsp;|&nbsp;
        <strong>IRA +8% qualifying: Biosimilar ASP ≤ Reference ASP 자동 검증</strong> &nbsp;|&nbsp;
        제품 매핑: Short Description 키워드 기반<br>
        <span style="color:#C62828;font-weight:bold;">━━</span> Originator &nbsp;|&nbsp;
        <span style="color:#1565C0;font-weight:bold;">╌╌</span> Samsung Bioepis &nbsp;|&nbsp;
        <span style="color:#78909C;">┄┄</span> 기타 Biosimilar &nbsp;|&nbsp;
        <span style="background:#1565C0;color:white;padding:1px 6px;border-radius:3px;font-size:10px;">★ SB</span> Samsung Bioepis 제품 있는 Molecule
      </div>
      {diag_html}
      {"<div style='padding:0 28px;'>" + val_html + new_code_html + "</div>" if val_html or new_code_html else ""}
    """

    def render_mol_section(mol_name, mol_data, cid_idx):
        quarters     = mol_data["quarters"]
        products     = mol_data["products"]
        display_dose = mol_data["display_dose"]
        has_sb       = mol_data["has_sb"]
        cid          = chart_cids[cid_idx]
        orig         = products[0]

        header_color = "#1565C0" if has_sb else "#37474F"
        sb_label     = ' <span style="background:white;color:#1565C0;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:bold;">★ SB</span>' if has_sb else ""

        q_headers = "".join(
            f'<th style="text-align:right;padding:6px 10px;white-space:nowrap;font-size:11px;">{q}</th>'
            for q in quarters)

        rows_html = ""
        for prod in products:
            if prod["is_originator"]:
                row_bg = "#FFF3F3"
                badge  = '<span style="background:#C62828;color:white;padding:1px 6px;border-radius:3px;font-size:10px;margin-left:5px;">Originator</span>'
            elif prod["is_sb"]:
                row_bg = "#E3F0FF"
                badge  = '<span style="background:#1565C0;color:white;padding:1px 6px;border-radius:3px;font-size:10px;margin-left:5px;">Samsung Bioepis</span>'
            else:
                row_bg = "white"
                badge  = ""

            fda_badge = (' <span style="color:#aaa;font-size:10px;">(FDA 미확인)</span>'
                         if prod.get("fda_status") == "unknown" else "")

            name_cell = (
                f'<td style="padding:7px 10px;background:{row_bg};border-bottom:1px solid #F0F0F0;min-width:200px;">'
                f'<strong style="font-size:12px;">{prod["brand"]}</strong>{badge}{fda_badge}<br>'
                f'<span style="font-size:10px;color:#555;">{prod["suffix"]}</span><br>'
                f'<span style="font-size:10px;color:#888;">{prod["company"]}</span>'
                f'</td>'
            )

            mult = prod["mult"]
            price_cells = ""
            prev_asp = None
            for r in prod["asp_results"]:
                if r is None:
                    price_cells += f'<td style="text-align:right;padding:7px 10px;background:{row_bg};color:#ddd;border-bottom:1px solid #F0F0F0;font-size:11px;">—</td>'
                else:
                    asp_conv = round(r["asp"] * mult, 2)
                    pl_conv  = round(r["payment_limit"] * mult, 2)
                    addon    = r["addon_pct"]
                    ira_q    = r.get("ira_qualifying", False)

                    addon_badge = (
                        '<span style="background:#F3E5F5;color:#7B1FA2;padding:0 3px;border-radius:2px;font-size:9px;font-weight:bold;">+8%IRA</span>'
                        if addon == 8 else
                        '<span style="font-size:9px;color:#bbb;">+6%</span>'
                    )
                    arrow = ""
                    if prev_asp is not None:
                        pct = (asp_conv - prev_asp) / prev_asp * 100
                        if pct > 0.5:
                            arrow = f'<span style="font-size:9px;color:#C62828;">▲{pct:.1f}%</span> '
                        elif pct < -0.5:
                            arrow = f'<span style="font-size:9px;color:#1565C0;">▼{abs(pct):.1f}%</span> '
                    price_cells += (
                        f'<td style="text-align:right;padding:7px 10px;background:{row_bg};border-bottom:1px solid #F0F0F0;">'
                        f'{arrow}<strong style="font-size:12px;">${asp_conv:,.2f}</strong> {addon_badge}<br>'
                        f'<span style="font-size:9px;color:#aaa;">PL ${pl_conv:,.2f}</span>'
                        f'</td>'
                    )
                    prev_asp = asp_conv
            rows_html += f"<tr>{name_cell}{price_cells}</tr>"

        return f"""
        <div style="margin:18px 0;border:1px solid #D0D7E3;border-radius:8px;overflow:hidden;">
          <div style="background:{header_color};color:white;padding:11px 18px;">
            <h3 style="margin:0;font-size:14px;">{mol_name}{sb_label}</h3>
            <span style="font-size:11px;opacity:.85;">
              Originator: {orig["brand"]} ({orig["company"]}) &nbsp;|&nbsp;
              {mol_data["benefit"]} &nbsp;|&nbsp; 기준: {display_dose}
            </span>
          </div>
          <div style="overflow-x:auto;">
            <table style="width:100%;border-collapse:collapse;font-size:12px;">
              <thead>
                <tr style="background:#F0F4FA;border-bottom:2px solid #D0D7E3;">
                  <th style="text-align:left;padding:7px 10px;min-width:200px;">Brand / INN suffix / Company</th>
                  {q_headers}
                </tr>
              </thead>
              <tbody>{rows_html}</tbody>
            </table>
          </div>
          <div style="background:#F8F9FA;padding:14px 18px;border-top:1px solid #E0E0E0;">
            <img src="cid:{cid}" style="width:100%;border-radius:6px;" />
          </div>
        </div>"""

    # Section 1: SB 제품 있는 molecule
    if sb_mols:
        html += """
        <div style="margin:20px 28px 8px;padding:8px 14px;background:#E3F0FF;border-left:4px solid #1565C0;border-radius:4px;">
          <strong style="color:#1565C0;">★ Section 1 — Samsung Bioepis 제품 포함 Molecule</strong>
        </div>"""
        for mol_name, mol_data in sb_mols:
            cid_idx = list(asp_data.keys()).index(mol_name)
            html += render_mol_section(mol_name, mol_data, cid_idx)

    # Section 2: SB 제품 없는 molecule (시장 검증용)
    if nosb_mols:
        html += """
        <div style="margin:28px 28px 8px;padding:8px 14px;background:#ECEFF1;border-left:4px solid #546E7A;border-radius:4px;">
          <strong style="color:#546E7A;">Section 2 — 시장 검증용 Molecule (Samsung Bioepis 제품 없음)</strong>
        </div>"""
        for mol_name, mol_data in nosb_mols:
            cid_idx = list(asp_data.keys()).index(mol_name)
            html += render_mol_section(mol_name, mol_data, cid_idx)

    html += f"""
      <div style="margin:20px 0;border:1px solid #E0E0E0;border-radius:8px;overflow:hidden;">
        <div style="background:#546E7A;color:white;padding:10px 20px;">
          <strong>모니터링 제외 제품 현황</strong>
        </div>
        <div style="padding:12px 18px;background:#FAFAFA;font-size:12px;">
          <ul style="margin:0;padding-left:18px;line-height:1.9;">{excluded_li}</ul>
        </div>
      </div>
      <div style="background:#F0F4FA;padding:12px 28px;border-radius:0 0 10px 10px;font-size:11px;color:#888;text-align:center;">
        Data Source: CMS Medicare Part B ASP Pricing Files (자동 수집) |
        IRA qualifying 자동 검증 포함 | Short Description 키워드 매핑 |
        {datetime.now().strftime("%Y-%m-%d %H:%M")} 생성
      </div>
    </div>"""
    return html


# ============================================================
# 7. 이메일 발송
# ============================================================
def send_email(html_body: str, charts: dict):
    msg = MIMEMultipart("related")
    today = datetime.now().strftime("%Y-%m-%d")
    msg["Subject"] = f"[Biosimilar ASP Monitor] {today} — SB Market Report 검증 (10 Molecules, 최대 12Q)"
    msg["From"] = EMAIL_CONFIG["sender"]
    msg["To"]   = ", ".join(EMAIL_CONFIG["recipients"])
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    for mol_name, png_bytes in charts.items():
        img = MIMEImage(png_bytes, "png")
        cid = f"chart_{mol_name.lower().replace(' ','_').replace('/','_')}"
        img.add_header("Content-ID", f"<{cid}>")
        img.add_header("Content-Disposition", "inline", filename=f"{cid}.png")
        msg.attach(img)
    print("\n[4/4] 이메일 발송 중...")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_CONFIG["sender"], EMAIL_CONFIG["app_password"])
        smtp.sendmail(EMAIL_CONFIG["sender"], EMAIL_CONFIG["recipients"], msg.as_bytes())
    print(f"✅ 발송 완료 -> {EMAIL_CONFIG['recipients']}")


# ============================================================
# 메인
# ============================================================
def main():
    print("=" * 65)
    print("  Biosimilar ASP Monitor | FINAL VERSION")
    print("  10 Molecules | 최대 12개 분기 | IRA qualifying 자동 검증")
    print("=" * 65)

    validation = validate_products()
    new_codes  = detect_new_hcpcs_codes()
    asp_data, all_diag = collect_asp_data(validation)

    print("\n[3/4] 그래프 생성 중...")
    charts, chart_cids = {}, []
    for mol_name, mol_data in asp_data.items():
        png = make_chart(mol_name, mol_data)
        cid = f"chart_{mol_name.lower().replace(' ','_').replace('/','_')}"
        charts[mol_name] = png
        chart_cids.append(cid)
        print(f"     -> {mol_name} 완료")

    html = build_email_html(asp_data, chart_cids, validation, new_codes, all_diag)
    send_email(html, charts)


if __name__ == "__main__":
    main()

# CSV 저장 (GitHub Pages용)
import csv, os

os.makedirs("data", exist_ok=True)
rows = []
for mol_name, mol_data in asp_data.items():
    for prod in mol_data["products"]:
        for i, r in enumerate(prod["asp_results"]):
            if r is not None:
                rows.append({
                    "quarter":    mol_data["quarters"][i],
                    "molecule":   mol_name,
                    "brand":      prod["brand"],
                    "company":    prod["company"],
                    "is_sb":      prod["is_sb"],
                    "is_orig":    prod["is_originator"],
                    "asp":        round(r["asp"] * prod["mult"], 2),
                    "addon_pct":  r["addon_pct"],
                    "ira":        r.get("ira_qualifying", False),
                    "hcpcs":      r.get("hcpcs", ""),
                })

with open("data/asp_data.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=rows[0].keys())
    writer.writeheader()
    writer.writerows(rows)

print("✅ data/asp_data.csv 저장 완료")
