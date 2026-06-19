"""
달콤수학 CS 챗봇 데이터 빌더
Excel raw-data → chatbot.html 생성
"""
import pandas as pd
import json
from pathlib import Path

EXCEL_PATH = Path(__file__).parent / "달콤교육_cs 로그.xlsx"
OUT_PATH   = Path(__file__).parent / "chatbot.html"

# ── 중분류 병합 맵 {대분류: {구_L2: 신_L2}} ───────────────────────────────────
MID_MERGE = {
    "수강": {
        "시청":       "강의 시청",
        "시청방법":   "강의 시청",
        "교재":       "교재·교구",
        "교구":       "교재·교구",
        "등업":       "등업·커뮤니티",
        "로드맵":     "로드맵·추천",
        "기간연장":   "기타",
        "기간중지":   "기타",
        "기간확인":   "기타",
        "연장":       "기타",
        "부분연장":   "기타",
        "부분환불":   "기타",
        "수강중지":   "기타",
        "배송":       "기타",
        "배송현황":   "기타",
        "배송확인":   "기타",
        "오픈":       "기타",
        "일정":       "기타",
        "재등록":     "기타",
        "전체환불":   "기타",
        "할인":       "기타",
        "구매링크":   "기타",
        "구입링크":   "기타",
    },
    "일정": {
        "오픈":   "강의 오픈",
        "일정":   "강의 오픈",
        "마감":   "기타",
        "확인":   "기타",
    },
    "환불": {
        "전체환불":   "전체 환불",
        "전액환불":   "전체 환불",
        "환불":       "전체 환불",
        "부분환불":   "부분 환불",
        "교재":       "교재 교환·환불",
        "기간확인":   "기타",
        "시청방법":   "기타",
    },
    "기타": {
        "로드맵":  "로드맵·추천",
        "교구":    "기타",
        "밴드":    "기타",
        "시청":    "기타",
    },
    "구입": {
        "교재":       "교재·워크지",
        "구입링크":   "구입 링크",
        "오픈":       "기타",
        "할인":       "기타",
    },
    "배송": {
        "배송현황":   "배송 문의",
    },
}

# ── 경로별 답변 ──────────────────────────────────────────────────────────────
SPECIFIC = {
    # 수강
    ("수강","강의 시청","방법"):         "강의는 달콤수학 앱 또는 홈페이지에 로그인한 뒤 [마이페이지 → 수강중인 강의]에서 시청하실 수 있어요.",
    ("수강","강의 시청","다시보기"):     "라이브 강의 다시보기는 방송 종료 후 1~2 영업일 내에 강의실에 업로드됩니다.",
    ("수강","강의 시청","종료일"):       "수강 종료일은 [마이페이지 → 수강내역]에서 확인하실 수 있어요.",
    ("수강","강의 시청","일정"):         "강의 편성 일정은 홈페이지 강의 페이지 또는 공식 SNS에서 확인하실 수 있어요.",
    ("수강","강의 시청","기타"):         "강의 시청 관련 기타 문의입니다. 상담원 연결을 이용해 주세요.",
    ("수강","재수강","가능"):            "재수강 신청은 [마이페이지 → 수강내역]에서 가능해요. 재수강 할인 혜택도 함께 확인해 보세요! 🎉",
    ("수강","재수강","불가능"):          "해당 강의는 현재 재수강이 어렵습니다. 자세한 안내는 상담원 연결을 이용해 주세요.",
    ("수강","재수강","금액"):            "재수강 금액은 홈페이지 강의 페이지에서 확인하실 수 있어요.",
    ("수강","재수강","방법"):            "재수강은 [마이페이지 → 수강내역 → 재수강 신청]에서 진행하실 수 있어요.",
    ("수강","재수강","할인불가"):        "죄송합니다, 해당 강의는 재수강 할인 적용이 어렵습니다.",
    ("수강","교재·교구","다운로드"):     "교재(워크지)는 강의실 내 [학습 자료] 탭에서 다운로드하실 수 있어요.",
    ("수강","교재·교구","배송"):         "실물 교재 배송 문의입니다. 주문 내역에서 배송 상태를 확인해 주세요.",
    ("수강","교재·교구","교환"):         "교재 교환은 수령 후 7일 이내 가능합니다. 불량 사진과 함께 상담원 연결을 이용해 주세요.",
    ("수강","교재·교구","추가구매"):     "교재 추가 구매는 홈페이지 스토어 또는 공구 링크에서 가능해요.",
    ("수강","교재·교구","추가구입"):     "교재 추가 구입은 홈페이지 스토어 또는 공구 링크에서 가능해요.",
    ("수강","교재·교구","자료 재발송"):  "자료 재발송 요청은 상담원 연결을 통해 처리해 드립니다.",
    ("수강","교재·교구","해외배송"):     "해외 배송 가능 여부는 상담원 연결을 통해 확인해 주세요.",
    ("수강","교재·교구","기타"):         "교재·교구 관련 기타 문의는 상담원 연결을 이용해 주세요.",
    ("수강","등업·커뮤니티","까페"):     "카페 등업은 스터디 활동 완료 후 운영진이 확인하여 처리합니다.",
    ("수강","등업·커뮤니티","단톡"):     "단톡방 초대는 수강 신청 완료 후 안내 메시지를 통해 진행됩니다.",
    ("수강","등업·커뮤니티","단톡방"):   "단톡방 초대는 수강 신청 완료 후 안내 메시지를 통해 진행됩니다.",
    ("수강","등업·커뮤니티","밴드"):     "밴드 가입 안내는 수강 신청 완료 후 발송되는 안내 메시지를 확인해 주세요.",
    ("수강","등업·커뮤니티","스윗맘"):   "스윗맘 스터디 관련 문의입니다. 신청 및 등업은 안내 페이지를 확인해 주세요.",
    ("수강","등업·커뮤니티","카톡"):     "카톡 알림 관련 문의는 상담원 연결을 통해 도움받으실 수 있어요.",
    ("수강","로드맵·추천","가능"):       "아이 연령과 수준에 맞는 강의 로드맵 추천은 상담원 연결을 통해 상세히 안내해 드려요.",
    ("수강","로드맵·추천","기타"):       "강의 로드맵 관련 기타 문의는 상담원 연결을 이용해 주세요.",
    # 환불
    ("환불","전체 환불","가능"):         "전액 환불은 수강 시작 전 가능합니다. 환불 처리는 영업일 기준 3~5일 소요됩니다.",
    ("환불","전체 환불","불가능"):       "수강 진행 후에는 전액 환불이 어렵습니다. 부분 환불 가능 여부는 상담원에게 문의해 주세요.",
    ("환불","부분 환불","가능"):         "부분 환불은 잔여 수강 기간에 따라 산정됩니다. 정확한 금액은 상담원 연결을 이용해 주세요.",
    ("환불","교재 교환·환불","교환"):    "불량 교재 교환은 수령 후 7일 이내 가능합니다. 사진과 함께 상담원 연결을 이용해 주세요.",
    ("환불","교재 교환·환불","기타"):    "교재 환불/교환 기타 문의는 상담원 연결을 이용해 주세요.",
    # 일정
    ("일정","강의 오픈","예정"):         "강의/공구 오픈 예정은 달콤수학 공식 SNS와 홈페이지 공지에서 가장 빠르게 확인하실 수 있어요! 📢",
    ("일정","강의 오픈","마감"):         "강의 신청 마감 일정은 해당 강의 페이지에서 확인하실 수 있어요.",
    ("일정","강의 오픈","확인"):         "강의 일정은 홈페이지 이벤트/강의 페이지에서 확인하실 수 있어요.",
    # 구입
    ("구입","교재·워크지","다운로드"):   "워크지 다운로드는 강의실 내 [학습 자료] 탭에서 가능해요.",
    ("구입","구입 링크","구입링크"):     "공구 제품 구입 링크는 달콤수학 카페 또는 밴드 공지에서 확인하실 수 있어요.",
    # 배송
    ("배송","배송 문의","확인"):         "배송 현황은 주문 완료 문자의 운송장 번호로 조회하실 수 있어요.",
    # 기타
    ("기타","로드맵·추천","가능"):       "아이 연령과 수준에 맞는 강의 추천은 상담원 연결을 통해 안내해 드릴게요.",
    ("기타","로드맵·추천","기타"):       "강의 추천 관련 기타 문의는 상담원 연결을 이용해 주세요.",
}

SUB_FALLBACK = {
    "가능":     "{l2}이 가능합니다. 상세 절차는 마이페이지 또는 홈페이지를 확인해 주세요.",
    "불가":     "죄송합니다. 현재 {l2}은 처리가 어렵습니다. 상담원에게 문의해 주세요.",
    "불가능":   "죄송합니다. 현재 {l2}은 처리가 어렵습니다. 상담원에게 문의해 주세요.",
    "방법":     "{l2} 방법은 로그인 후 마이페이지에서 확인하실 수 있어요.",
    "확인":     "{l2} 확인은 마이페이지에서 바로 조회하실 수 있어요.",
    "현황":     "{l2} 현황은 마이페이지 또는 주문 내역에서 확인하실 수 있어요.",
    "금액":     "{l2} 금액은 홈페이지 강의 페이지에서 확인하실 수 있어요.",
    "예정":     "{l2} 예정 일정은 공식 SNS와 홈페이지 공지에서 확인하실 수 있어요.",
    "마감":     "{l2} 마감 일정은 해당 강의 페이지에서 확인해 주세요.",
    "다운로드": "다운로드는 강의실 [학습 자료] 탭에서 가능해요.",
}

def get_answer(l1, l2, l3):
    if (l1, l2, l3) in SPECIFIC:
        return SPECIFIC[(l1, l2, l3)]
    tmpl = SUB_FALLBACK.get(l3)
    if tmpl:
        return tmpl.format(l1=l1, l2=l2, l3=l3)
    return f"{l2} 관련 문의입니다. 상세한 안내를 위해 상담원이 도움드리겠습니다."

# ── 데이터 로드 및 정제 ──────────────────────────────────────────────────────
def load_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name="raw-data")

    def clean(s):
        if pd.isna(s): return None
        s = str(s).strip().replace("젠체환불", "전체환불").replace("ㅅ시청", "시청")
        return s or None

    df["L1"] = df["문의유형"].apply(clean).replace("기간", "수강")
    df["L2"] = df["태그1"].apply(clean)
    df["L3"] = df["태그2"].apply(clean)
    df["S"]  = df["질문요약"].apply(lambda x: str(x).strip() if not pd.isna(x) else None)

    # MID_MERGE 적용
    def apply_merge(row):
        l1, l2 = row["L1"], row["L2"]
        if l1 and l2 and l1 in MID_MERGE:
            return MID_MERGE[l1].get(l2, l2)
        return l2

    df["L2"] = df.apply(apply_merge, axis=1)

    # 5건 미만 → 기타
    counts = df.groupby(["L1", "L2"]).size()
    def mark_small(row):
        if row["L2"] and row["L2"] != "기타":
            cnt = counts.get((row["L1"], row["L2"]), 0)
            if cnt < 5:
                return "기타"
        return row["L2"]

    df["L2"] = df.apply(mark_small, axis=1)
    return df

# ── 최대 6개 제한 ────────────────────────────────────────────────────────────
CAT_ORDER = ["수강", "일정", "환불", "기타", "구입", "배송"]

def enforce_max6(df):
    df = df.copy()
    for l1 in CAT_ORDER:
        mask_l1 = df["L1"] == l1
        d1 = df[mask_l1 & (df["L2"] != "기타")]
        counts = d1.groupby("L2").size().sort_values(ascending=False)
        if len(counts) > 5:
            keep = set(counts.head(5).index)
            overflow_mask = mask_l1 & (~df["L2"].isin(keep)) & (df["L2"] != "기타")
            df.loc[overflow_mask, "L2"] = "기타"
    return df

# ── 트리 빌드 ────────────────────────────────────────────────────────────────
def build_tree(df):
    tree = {}
    for l1 in CAT_ORDER:
        d1 = df[df["L1"] == l1]
        if d1.empty: continue
        mid_order = [m for m in d1["L2"].dropna().unique() if m != "기타"]
        mid_order = sorted(mid_order) + (["기타"] if "기타" in d1["L2"].values else [])
        tree[l1] = {}
        for l2 in mid_order:
            d2 = d1[d1["L2"] == l2]
            subs = sorted(d2["L3"].dropna().unique())
            if not subs:
                subs = ["기타"]
            tree[l1][l2] = {}
            for l3 in subs:
                d3 = d2[d2["L3"] == l3] if l3 in d2["L3"].values else d2
                samples = d3["S"].dropna().unique().tolist()[:4]
                tree[l1][l2][l3] = {"answer": get_answer(l1, l2, l3), "samples": samples}
    return tree

# ── 키워드 인덱스 ────────────────────────────────────────────────────────────
def build_keywords(df):
    seen, result = set(), []
    for _, row in df.iterrows():
        if not row["S"]: continue
        path = [x for x in [row["L1"], row["L2"], row["L3"]] if x]
        key  = tuple(path + [row["S"]])
        if key in seen or len(path) < 2: continue
        seen.add(key)
        result.append({"s": row["S"], "p": path})
    return result[:800]

# ── HTML 템플릿 ───────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>달콤수학 상담 챗봇</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Apple SD Gothic Neo','Noto Sans KR',sans-serif;background:#b2c7d9;min-height:100vh;display:flex;justify-content:center;align-items:center;padding:20px}
.wrap{width:100%;max-width:420px;height:720px;background:#b2c7d9;display:flex;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.25)}
/* 헤더 */
.hdr{background:#3a1f1f;color:#fff;padding:14px 16px;display:flex;align-items:center;gap:10px;flex-shrink:0}
.hdr-icon{width:38px;height:38px;border-radius:50%;background:#fff9e6;display:flex;align-items:center;justify-content:center;overflow:hidden;flex-shrink:0}
.hdr-icon img{width:30px;height:30px;object-fit:contain}
.hdr-text{font-size:15px;font-weight:700;letter-spacing:-.3px}
.hdr-sub{font-size:11px;opacity:.7;margin-top:1px}
/* 메시지 영역 */
.msgs{flex:1;overflow-y:auto;padding:16px 12px;display:flex;flex-direction:column;gap:10px;scroll-behavior:smooth}
/* 봇 메시지 */
.bot-row{display:flex;align-items:flex-start;gap:8px}
.bot-ava{width:34px;height:34px;border-radius:50%;background:#fff9e6;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px;overflow:hidden}
.bot-ava img{width:28px;height:28px;object-fit:contain}
.bot-ava .fallback{font-size:18px;line-height:1}
.bot-col{display:flex;flex-direction:column;gap:6px;max-width:285px}
.bot-name{font-size:11px;color:#555;margin-bottom:2px}
.bubble{background:#fff;border-radius:0 14px 14px 14px;padding:10px 14px;font-size:14px;line-height:1.6;color:#222;word-break:keep-all}
.bubble.wide{max-width:100%}
/* 유저 메시지 */
.user-row{display:flex;justify-content:flex-end}
.user-bubble{background:#fee500;border-radius:14px 0 14px 14px;padding:10px 14px;font-size:14px;line-height:1.6;max-width:240px;word-break:keep-all}
/* 버튼 그룹 */
.btns{display:flex;flex-direction:column;gap:5px}
.btns.row{flex-direction:row;flex-wrap:wrap;gap:6px}
.btn{background:#fff;border:1.5px solid #ddd;border-radius:10px;padding:9px 14px;font-size:13px;cursor:pointer;color:#333;text-align:left;transition:background .12s,border-color .12s;font-family:inherit;word-break:keep-all}
.btn:hover{background:#fffde7;border-color:#fee500}
.btn.yellow{background:#fee500;border-color:#fee500;font-weight:700}
.btn.yellow:hover{background:#ffd600}
.btn.agent{background:#3a1f1f;color:#fff;border-color:#3a1f1f;font-weight:700;text-align:center;width:100%}
.btn.agent:hover{background:#5c3030}
.btn.resolved{background:#e8f5e9;border-color:#66bb6a;color:#2e7d32;font-weight:600}
.btn.small{padding:7px 12px;font-size:12px}
/* 경로 / 태그 */
.tag{display:inline-block;background:#fff3e0;color:#e65100;border-radius:6px;padding:2px 7px;font-size:11px;margin-right:4px;margin-bottom:3px}
.warn{background:#fff8e1;border:1.5px solid #ffca28;border-radius:12px;padding:12px 14px;font-size:13px;color:#5d4037;line-height:1.6}
.path-label{font-size:11px;color:#888;display:block;margin-bottom:4px}
.path-label span{color:#c8860a;font-weight:600}
/* 입력창 */
.inp-bar{padding:10px 12px;background:#fff;border-top:1px solid #ddd;display:none;gap:8px;align-items:center;flex-shrink:0}
.inp-bar.show{display:flex}
.inp-bar input{flex:1;border:1.5px solid #ddd;border-radius:20px;padding:9px 14px;font-size:14px;outline:none;font-family:inherit}
.inp-bar input:focus{border-color:#fee500}
.inp-bar button{background:#fee500;border:none;border-radius:50%;width:38px;height:38px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.inp-bar button:hover{background:#ffd600}
.msgs::-webkit-scrollbar{width:4px}
.msgs::-webkit-scrollbar-thumb{background:#ccc;border-radius:2px}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <div class="hdr-icon">
      <img src="dalcong.png" alt="달콩" onerror="this.style.display='none';this.parentNode.innerHTML='<span class=\'fallback\'>🍯</span>'">
    </div>
    <div>
      <div class="hdr-text">달콤수학 상담 챗봇</div>
      <div class="hdr-sub">평일 10:00 – 18:00 운영</div>
    </div>
  </div>
  <div class="msgs" id="msgs"></div>
  <div class="inp-bar" id="inpBar">
    <input id="inpText" placeholder="질문을 입력하세요…" maxlength="200" />
    <button onclick="submitText()">↑</button>
  </div>
</div>

<script>
const DATA = __CHATBOT_DATA__;

// ── 강의 분류 트리 (하드코딩) ────────────────────────────────────────────────
const CLASS_TREE = {
  "프로젝트": [
    "유아사고력 1단계 프로젝트",
    "유아사고력 2단계 프로젝트",
    "초등사고력 1단계 프로젝트",
    "초등사고력 2단계 프로젝트",
    "초등사고력 3단계 프로젝트"
  ],
  "특강": [
    "수배열판 특강",
    "직산 특강",
    "소마큐브 특강",
    "분수소수 특강",
    "곱셈나눗셈 특강",
    "로드맵 특강"
  ]
};
const CLASS_ANSWERS = {
  "프로젝트": "해당 강의 일정은 자사몰(홈페이지)에서 확인하실 수 있어요. 자세한 안내가 필요하면 상담원 연결을 이용해 주세요.",
  "특강": "특강 목록과 일정은 자사몰에서 확인하실 수 있어요. 자세한 안내가 필요하면 상담원 연결을 이용해 주세요."
};
// 강의 일정 관련 자유 텍스트 감지 키워드
const SCHEDULE_KW = [
  "강의 일정","수업 언제","언제 열","오픈 언제","특강 언제","프로젝트 언제",
  "수배열","직산","소마큐브","분수소수","곱셈나눗셈","로드맵 특강",
  "유아사고력","초등사고력","1단계","2단계","3단계"
];

// ── 금지 키워드 ──────────────────────────────────────────────────────────────
const FORBIDDEN = [
  "사업자번호","사업자등록","대표자 이름","대표자 연락처",
  "수강생 이름","학생 연락처","수강생 명단","수강생 주소","수강 인원",
  "운영 비용","수수료","마진","원가","강사 전화","강사 개인 연락",
  "미공개 커리큘럼","강의 단가","내부 가격","경쟁사","타사 비교","vs 달콤",
  "환불 계산","얼마 돌려","고소","소송","법적 조치","내용증명",
  "미공개 과정","출시 예정 강의"
];

// ── 상태 ────────────────────────────────────────────────────────────────────
let mode = null;       // 'tree' | 'text' | 'keyword'
let depth = 0;
let path  = [];
let txtInteractions = 0;

// ── DOM ──────────────────────────────────────────────────────────────────────
const $msgs = document.getElementById('msgs');
const $bar  = document.getElementById('inpBar');
const $inp  = document.getElementById('inpText');

function scroll(){ $msgs.scrollTop = $msgs.scrollHeight; }
function esc(s){ const d=document.createElement('div'); d.textContent=s; return d.innerHTML; }

// 봇 아바타 HTML 생성 (이미지 없으면 이모지 폴백)
function avatarHTML(){
  return `<img src="dalcong.png" alt="달콩" onerror="this.style.display='none';this.parentNode.innerHTML='<span class=\\'fallback\\'>🍯</span>'">`;
}

function botRow(html, extra=''){
  const row = document.createElement('div');
  row.className = 'bot-row';
  const ava = document.createElement('div');
  ava.className = 'bot-ava';
  ava.innerHTML = avatarHTML();
  const col = document.createElement('div');
  col.className = 'bot-col';
  col.innerHTML = `<div class="bot-name">달콤수학 CS봇</div><div class="bubble ${extra}">${html}</div>`;
  row.appendChild(ava);
  row.appendChild(col);
  $msgs.appendChild(row);
  scroll();
}

function userRow(text){
  const row = document.createElement('div');
  row.className = 'user-row';
  const bub = document.createElement('div');
  bub.className = 'user-bubble';
  bub.textContent = text;  // textContent로 안전하게 설정
  row.appendChild(bub);
  $msgs.appendChild(row);
  scroll();
}

// 버튼 그룹 — DOM 직접 생성으로 클로저·파싱 버그 방지
function buttonsRow(buttons, rowLayout){
  const rowDiv = document.createElement('div');
  rowDiv.className = 'bot-row';

  const spacer = document.createElement('div');
  spacer.style.cssText = 'width:34px;flex-shrink:0';

  const col = document.createElement('div');
  col.className = 'bot-col';
  col.style.maxWidth = '340px';

  const wrap = document.createElement('div');
  wrap.className = rowLayout ? 'btns row' : 'btns';

  buttons.forEach(function(b){
    const btn = document.createElement('button');
    btn.className = 'btn ' + (b.cls || '');
    btn.textContent = b.label;                 // textContent → 텍스트 잘림 방지
    const fn = b.fn;                           // 명시적 캡처
    btn.addEventListener('click', function(e){
      e.stopPropagation();
      fn();
    });
    wrap.appendChild(btn);
  });

  col.appendChild(wrap);
  rowDiv.appendChild(spacer);
  rowDiv.appendChild(col);
  $msgs.appendChild(rowDiv);
  scroll();
}

// ── 초기화 ──────────────────────────────────────────────────────────────────
function init(){
  botRow('안녕하세요! 달콤수학 상담 챗봇입니다 😊<br>어떻게 도와드릴까요?');
  buttonsRow([
    {label:'📋  메뉴에서 선택하기',  fn: function(){ startMode('tree'); }},
    {label:'✏️  직접 질문 입력하기',  fn: function(){ startMode('text'); }},
    {label:'🔍  키워드로 검색하기',   fn: function(){ startMode('keyword'); }},
  ]);
}

// ── 모드 선택 ────────────────────────────────────────────────────────────────
function startMode(m){
  mode = m; depth = 0; path = []; txtInteractions = 0;
  const label = {tree:'📋 메뉴에서 선택하기', text:'✏️ 직접 질문 입력하기', keyword:'🔍 키워드로 검색하기'}[m];
  userRow(label);

  if(m === 'tree'){
    $bar.classList.remove('show');
    showL1Menu();
  } else if(m === 'text'){
    $bar.classList.add('show');
    botRow('질문을 자유롭게 입력해 주세요.<br><span style="font-size:12px;color:#888">예) 강의 일정이 궁금해요.</span>');
    $inp.focus();
  } else {
    $bar.classList.add('show');
    botRow('검색할 키워드를 입력해 주세요.<br><span style="font-size:12px;color:#888">예) 환불, 배송, 재수강, 워크지</span>');
    $inp.focus();
  }
}

// ── 트리 모드 ────────────────────────────────────────────────────────────────
function showL1Menu(){
  botRow('어떤 유형의 문의인가요?');
  buttonsRow(
    Object.keys(DATA.tree).map(function(l1){
      return {label: l1, fn: function(){ selectL1(l1); }};
    }),
    true
  );
}

function selectL1(l1){
  depth = 1; path = [l1];
  userRow(l1);

  // 일정 → 강의 분류 트리 (CLASS_TREE)
  if(l1 === '일정'){
    showClassType();
    return;
  }

  const l2keys = Object.keys(DATA.tree[l1] || {});
  if(!l2keys.length){ depth = 3; showAnswer(); return; }
  botRow('<span class="path-label"><span>' + esc(l1) + '</span> 선택됨</span>세부 유형을 선택해 주세요.');
  buttonsRow(l2keys.map(function(l2){
    return {label: l2, fn: function(){ selectL2(l1, l2); }};
  }));
}

function selectL2(l1, l2){
  depth = 2; path = [l1, l2];
  userRow(l2);

  const l3keys = Object.keys(DATA.tree[l1][l2] || {});
  if(!l3keys.length){
    // L3 없음 → depth=3 처리로 강제 (규칙: 3단계 후 상담원 연결)
    depth = 3; showAnswer(); return;
  }
  botRow('<span class="path-label"><span>' + esc(l1) + ' › ' + esc(l2) + '</span></span>구체적인 내용을 선택해 주세요.');
  buttonsRow(l3keys.map(function(l3){
    return {label: l3, fn: function(){ selectL3(l1, l2, l3); }};
  }));
}

function selectL3(l1, l2, l3){
  depth = 3; path = [l1, l2, l3];
  userRow(l3);
  showAnswer();
}

function showAnswer(){
  const l1 = path[0], l2 = path[1], l3 = path[2];
  let node = null;
  if(l3 && DATA.tree[l1] && DATA.tree[l1][l2] && DATA.tree[l1][l2][l3]){
    node = DATA.tree[l1][l2][l3];
  } else if(l2 && DATA.tree[l1] && DATA.tree[l1][l2]){
    const subs = Object.values(DATA.tree[l1][l2]);
    if(subs.length === 1) node = subs[0];
  }
  const answer  = node ? node.answer  : (l1 + ' 관련 문의입니다. 상담원 연결을 통해 안내해 드릴게요.');
  const samples = node ? node.samples : [];

  let html = '<span class="path-label"><span>' + esc(path.join(' › ')) + '</span></span>' + esc(answer);
  if(samples.length){
    html += '<br><br><span style="font-size:12px;color:#888">유사 질문 예시:</span><br>';
    samples.forEach(function(s){ html += '<span class="tag">' + esc(s.slice(0,28)) + (s.length>28?'…':'') + '</span>'; });
  }
  botRow(html, 'wide');

  buttonsRow([
    {label:'✅  해결됐어요!',  cls:'resolved', fn: function(){ onResolved(); }},
    {label:'❌  아직 안 됐어요', fn: function(){ onNotResolved(); }},
  ]);
}

function onResolved(){
  userRow('해결됐어요!');
  botRow('도움이 됐다니 다행이에요 😊<br>달콤수학을 이용해 주셔서 감사합니다!');
  buttonsRow([{label:'🏠  처음으로 돌아가기', fn: function(){ restart(); }}]);
}

// 핵심 버그 수정: depth < 3이면 상담원 연결 바로 노출 금지
function onNotResolved(){
  userRow('아직 안 됐어요');
  if(depth >= 3){
    botRow('불편을 드려 죄송합니다.<br>상담원이 직접 도와드릴게요.');
    showAgentBtn();
  } else {
    botRow('더 자세히 찾아볼게요. 메뉴를 다시 선택하시거나 상담원에게 연결해 드릴게요.');
    buttonsRow([
      {label:'📋  메뉴 처음으로',  fn: function(){ userRow('메뉴 처음으로'); depth=0; path=[]; showL1Menu(); }},
      {label:'📞  상담원 연결하기', cls:'agent', fn: function(){ connectAgent(); }},
    ]);
  }
}

// ── 강의 분류 트리 (CLASS_TREE) ──────────────────────────────────────────────
function showClassType(){
  botRow('어떤 강의를 찾고 계신가요?');
  var btns = Object.keys(CLASS_TREE).map(function(type){
    return {label: type, fn: function(){ showClassList(type); }};
  });
  btns.push({label:'기타 일정 문의', fn: function(){
    depth = 3; path = ['일정','기타','일정'];
    userRow('기타 일정 문의');
    botRow('기타 일정 문의입니다.<br>강의/공구 오픈 일정은 달콤수학 공식 SNS와 홈페이지에서 가장 빠르게 확인하실 수 있어요! 📢');
    buttonsRow([
      {label:'✅  해결됐어요!',  cls:'resolved', fn: function(){ onResolved(); }},
      {label:'❌  아직 안 됐어요', fn: function(){ onNotResolved(); }},
    ]);
  }});
  buttonsRow(btns);
}

function showClassList(type){
  depth = 2; path = ['일정', type];
  userRow(type);
  botRow('<span class="path-label"><span>일정 › ' + esc(type) + '</span></span>' + type + ' 중 어떤 강의인가요?');
  buttonsRow(CLASS_TREE[type].map(function(course){
    return {label: course, fn: function(){ selectCourse(type, course); }};
  }));
}

function selectCourse(type, course){
  depth = 3; path = ['일정', type, course];
  userRow(course);
  const answer = CLASS_ANSWERS[type] || '해당 강의 정보는 자사몰에서 확인해 주세요.';
  botRow('<span class="path-label"><span>일정 › ' + esc(type) + ' › ' + esc(course) + '</span></span>' + esc(answer));
  buttonsRow([
    {label:'✅  해결됐어요!',  cls:'resolved', fn: function(){ onResolved(); }},
    {label:'❌  아직 안 됐어요', fn: function(){ onNotResolved(); }},
  ]);
}

// ── 텍스트 입력 ──────────────────────────────────────────────────────────────
function submitText(){
  const text = $inp.value.trim();
  if(!text) return;
  $inp.value = '';
  if(checkForbidden(text)) return;
  userRow(text);
  txtInteractions++;
  if(mode === 'keyword') doKeywordSearch(text);
  else doTextMatch(text);
}
$inp.addEventListener('keydown', function(e){ if(e.key === 'Enter') submitText(); });

function checkForbidden(text){
  const hit = FORBIDDEN.find(function(k){ return text.includes(k); });
  if(!hit) return false;
  const rowDiv = document.createElement('div');
  rowDiv.className = 'bot-row';
  const ava = document.createElement('div'); ava.className = 'bot-ava'; ava.innerHTML = avatarHTML();
  const col = document.createElement('div'); col.className = 'bot-col'; col.style.maxWidth = '300px';
  col.innerHTML = '<div class="bot-name">달콤수학 CS봇</div><div class="warn">이 내용은 챗봇에서 답변드릴 수 없는 항목이에요.<br>상담이 필요하시면 아래 상담원 연결을 이용해 주세요.</div>';
  rowDiv.appendChild(ava); rowDiv.appendChild(col);
  $msgs.appendChild(rowDiv);
  showAgentBtn(true);
  scroll();
  return true;
}

function doTextMatch(text){
  // 강의 일정 키워드 감지
  if(SCHEDULE_KW.some(function(k){ return text.includes(k); })){
    botRow('강의 일정 관련 문의이군요! 어떤 강의를 찾고 계신가요?');
    showClassType();
    return;
  }

  const toks = text.replace(/[^가-힣\w]/g,' ').split(/\s+/).filter(function(t){ return t.length >= 2; });
  if(!toks.length){ offerMenu(); return; }

  var best = null, bestScore = 0;
  DATA.keywords.forEach(function(kw){
    var score = toks.reduce(function(s,t){ return s + (kw.s.includes(t) ? 1 : 0); }, 0);
    if(score > bestScore){ bestScore = score; best = kw; }
  });

  if(best && bestScore >= 1){
    path = best.p; depth = 3;
    const l1=best.p[0], l2=best.p[1], l3=best.p[2];
    const node = (l3 && DATA.tree[l1] && DATA.tree[l1][l2]) ? DATA.tree[l1][l2][l3] : null;
    const answer = node ? node.answer : (l1 + ' 관련 문의입니다. 상담원 연결을 통해 안내해 드릴게요.');
    let html = '<span class="path-label">매칭: <span>' + esc(best.p.join(' › ')) + '</span></span>';
    html += '<span style="font-size:12px;color:#888">"' + esc(best.s.slice(0,35)) + '…" 와 유사한 문의예요.</span><br><br>' + esc(answer);
    botRow(html, 'wide');
    buttonsRow([
      {label:'✅  해결됐어요!',  cls:'resolved', fn: function(){ onResolved(); }},
      {label:'❌  아직 안 됐어요', fn: function(){ txtNotResolved(); }},
    ]);
  } else {
    offerMenu();
  }
}

function doKeywordSearch(keyword){
  const toks = keyword.replace(/[^가-힣\w]/g,' ').split(/\s+/).filter(function(t){ return t.length >= 1; });
  const seen = {}, results = [];
  for(var i=0; i<DATA.keywords.length && results.length<5; i++){
    const kw = DATA.keywords[i];
    const match = toks.some(function(t){ return kw.s.includes(t) || kw.p.some(function(p){ return p.includes(t); }); });
    if(match){
      const key = kw.p.join('|');
      if(!seen[key]){ seen[key]=1; results.push(kw); }
    }
  }

  if(!results.length){
    botRow('관련 항목을 찾지 못했어요.<br>아래에서 메뉴를 직접 선택해 보시겠어요?');
    buttonsRow([
      {label:'📋  메뉴에서 선택하기', fn: function(){ startMode('tree'); }},
      {label:'✏️  직접 질문 입력하기', fn: function(){ startMode('text'); }},
    ]);
    if(txtInteractions >= 3) showAgentBtn();
    return;
  }

  botRow('"' + esc(keyword) + '" 관련 FAQ ' + results.length + '건을 찾았어요. 해당 항목을 선택해 주세요.');
  buttonsRow(results.map(function(r){
    const label = r.p.join(' › ') + ' — ' + r.s.slice(0,22) + (r.s.length>22?'…':'');
    return {label: label, cls:'small', fn: function(){ selectKeywordResult(r); }};
  }));
}

function selectKeywordResult(r){
  userRow(r.s.slice(0,30));
  path = r.p; depth = 3;
  showAnswer();
}

function offerMenu(){
  botRow('잘 이해하지 못했어요. 😅<br>아래 방식으로 다시 시도해 보시겠어요?');
  var btns = [{label:'📋  메뉴에서 선택하기', fn: function(){ startMode('tree'); }}];
  if(mode==='text') btns.push({label:'🔍  키워드로 검색하기', fn: function(){ startMode('keyword'); }});
  else              btns.push({label:'✏️  직접 질문 입력하기', fn: function(){ startMode('text'); }});
  if(txtInteractions >= 3) btns.push({label:'📞  상담원 연결하기', cls:'agent', fn: function(){ connectAgent(); }});
  buttonsRow(btns);
}

function txtNotResolved(){
  userRow('아직 안 됐어요');
  if(txtInteractions >= 3){
    botRow('불편을 드려 죄송합니다. 상담원이 직접 도와드릴게요.');
    showAgentBtn();
  } else {
    botRow('다른 방식으로 다시 시도해 보세요.');
    buttonsRow([
      {label:'📋  메뉴에서 선택하기', fn: function(){ startMode('tree'); }},
      {label:'🔍  키워드로 검색하기', fn: function(){ startMode('keyword'); }},
      {label:'📞  상담원 연결하기', cls:'agent', fn: function(){ connectAgent(); }},
    ]);
  }
}

// ── 상담원 연결 ──────────────────────────────────────────────────────────────
function showAgentBtn(manual){
  const rowDiv = document.createElement('div');
  rowDiv.className = 'bot-row';
  const spacer = document.createElement('div'); spacer.style.cssText = 'width:34px;flex-shrink:0';
  const col = document.createElement('div'); col.className = 'bot-col'; col.style.maxWidth = '300px';
  const btn = document.createElement('button');
  btn.className = 'btn agent';
  btn.textContent = manual ? '📞  상담원 연결하기 (직접 클릭)' : '📞  상담원 연결하기';
  btn.addEventListener('click', function(){ connectAgent(); });
  col.appendChild(btn);
  rowDiv.appendChild(spacer); rowDiv.appendChild(col);
  $msgs.appendChild(rowDiv);
  scroll();
}

function connectAgent(){
  userRow('상담원 연결하기');
  botRow('상담원과 연결 중입니다… ⏳<br><span style="font-size:12px;color:#888">(데모 버전: 실제 카카오톡 채널 상담 연결은 추후 적용됩니다)</span>');
  buttonsRow([{label:'🏠  처음으로 돌아가기', fn: function(){ restart(); }}]);
}

function restart(){
  $msgs.innerHTML = '';
  $bar.classList.remove('show');
  mode = null; depth = 0; path = []; txtInteractions = 0;
  init();
}

init();
</script>
</body>
</html>
"""

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    print("📊 Excel 로딩 중…")
    df = load_data()
    print(f"   총 {len(df)}행 로드됨")

    print("🔀 중분류 병합 및 최대 6개 제한 적용 중…")
    df = enforce_max6(df)

    print("🌳 트리 구조 빌드 중…")
    tree = build_tree(df)
    total = sum(len(v3) for v1 in tree.values() for v3 in v1.values())
    print(f"   대분류 {len(tree)}개, 총 경로 {total}개")
    for l1, l2dict in tree.items():
        print(f"   [{l1}] 중분류 {len(l2dict)}개: {list(l2dict.keys())}")

    print("🔑 키워드 인덱스 빌드 중…")
    keywords = build_keywords(df)
    print(f"   키워드 {len(keywords)}개")

    data = {"tree": tree, "keywords": keywords}
    json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    html = HTML_TEMPLATE.replace("__CHATBOT_DATA__", json_str)
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"✅  생성 완료: {OUT_PATH}")

if __name__ == "__main__":
    main()
