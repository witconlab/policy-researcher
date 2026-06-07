import streamlit as st
import google.generativeai as genai
import os
import json
import pdfplumber
import requests
from bs4 import BeautifulSoup
from pathlib import Path

st.set_page_config(
    page_title="정책 리서처",
    page_icon="📚",
    layout="wide",
)

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
.summary-card {
    background: #f8fdf8;
    border-left: 4px solid #2E7D32;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 12px;
}
.summary-card h4 { margin: 0 0 8px 0; color: #1B5E20; font-size: 0.95rem; }
.summary-card ul { margin: 0; padding-left: 18px; color: #333; font-size: 0.88rem; line-height: 1.7; }
.metric-box {
    background: #E8F5E9;
    border-radius: 10px;
    padding: 18px;
    text-align: center;
    margin-bottom: 8px;
}
.metric-box .num { font-size: 2rem; font-weight: 700; color: #2E7D32; }
.metric-box .label { font-size: 0.82rem; color: #555; margin-top: 4px; }
.region-bar {
    background: #fff;
    border: 1px solid #C8E6C9;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 8px;
}
.region-bar .region-name { font-weight: 600; color: #1B5E20; margin-bottom: 6px; font-size: 0.9rem; }
.bar-wrap { background: #E8F5E9; border-radius: 4px; height: 10px; }
.bar-fill { background: #2E7D32; border-radius: 4px; height: 10px; }
</style>
""", unsafe_allow_html=True)

# ── API 설정 ──────────────────────────────────────────────────
GOOGLE_API_KEY = st.secrets.get("GOOGLE_API_KEY", os.environ.get("GOOGLE_API_KEY", ""))
if not GOOGLE_API_KEY:
    st.error("Google API 키가 설정되지 않았습니다.")
    st.stop()

genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

POLICIES_DIR = Path("policies")


# ── PDF 로드 ──────────────────────────────────────────────────
@st.cache_resource(show_spinner="PDF 자료를 읽는 중...")
def load_policy_documents(policy_name: str) -> dict[str, str]:
    docs = {}
    policy_path = POLICIES_DIR / policy_name
    for pdf_file in sorted(policy_path.glob("**/*.pdf")) + sorted(policy_path.glob("**/*.PDF")):
        try:
            parts = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        parts.append(t)
            if parts:
                docs[pdf_file.name] = "\n".join(parts)
        except Exception:
            pass
    return docs


def get_relevant_chunks(query: str, docs: dict[str, str], max_chars: int = 60000) -> str:
    keywords = [w for w in query.lower().split() if len(w) > 1]
    scored = []
    for filename, text in docs.items():
        score = sum(text.lower().count(kw) for kw in keywords)
        if score > 0:
            scored.append((score, filename, text))
    scored.sort(reverse=True)
    if not scored:
        scored = [(0, fn, tx) for fn, tx in list(docs.items())[:5]]
    parts, total = [], 0
    for _, filename, text in scored:
        chunk = f"[출처: {filename}]\n{text[:8000]}\n"
        if total + len(chunk) > max_chars:
            break
        parts.append(chunk)
        total += len(chunk)
    return "\n---\n".join(parts)


def ask_gemini(query: str, context: str, history: list) -> str:
    system = f"""당신은 정책 문서를 분석하는 전문 리서처입니다.
아래 문서들을 바탕으로 질문에 답하세요.
- 답변은 한국어로 작성하세요.
- 근거 문서명을 반드시 인용하세요. 예: ([출처: 파일명.pdf])
- 문서에 없는 내용은 "자료에서 확인되지 않습니다"라고 답하세요.
- 비교 질문에는 표 형식을 활용하세요.

=== 참고 문서 ===
{context}
"""
    chat_history = []
    for msg in history[:-1]:
        role = "user" if msg["role"] == "user" else "model"
        chat_history.append({"role": role, "parts": [msg["content"]]})
    chat = model.start_chat(history=chat_history)
    return chat.send_message(f"{system}\n\n질문: {query}").text


@st.cache_data(show_spinner=False)
def generate_summary_cards(policy_name: str, doc_texts: str) -> list:
    prompt = f"""아래는 '{policy_name}' 관련 정책 문서들입니다.
지역 또는 연구 기관별로 핵심 내용을 요약해주세요.

반드시 아래 JSON 배열 형식으로만 응답하세요 (다른 텍스트 없이):
[
  {{
    "title": "지역/기관명 + 문서 제목 (30자 이내)",
    "points": ["핵심 내용 1", "핵심 내용 2", "핵심 내용 3"],
    "keyword": "핵심 키워드"
  }}
]

최대 8개 카드, 각 포인트는 한 문장으로 간결하게.

문서:
{doc_texts[:40000]}
"""
    try:
        resp = model.generate_content(prompt)
        text = resp.text.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def generate_infographic_data(policy_name: str, doc_texts: str) -> dict:
    prompt = f"""아래 정책 문서에서 인포그래픽에 활용할 수치와 비교 데이터를 추출하세요.

반드시 아래 JSON 형식으로만 응답하세요:
{{
  "metrics": [
    {{"label": "지표명", "value": "숫자/값", "unit": "단위(명/개/% 등)", "source": "출처문서명"}},
    ...
  ],
  "regions": [
    {{"name": "지역명", "approach": "인정체계 접근방식 한 줄", "score": 75}},
    ...
  ],
  "timeline": [
    {{"year": "2018", "event": "주요 사건/정책 한 줄"}},
    ...
  ],
  "key_issues": ["주요 과제 1", "주요 과제 2", "주요 과제 3", "주요 과제 4"]
}}

score는 제도화 수준을 0~100으로 추정. metrics 최대 6개, regions 최대 6개, timeline 최대 6개.

문서:
{doc_texts[:40000]}
"""
    try:
        resp = model.generate_content(prompt)
        text = resp.text.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return {}


def fetch_url_text(url: str) -> tuple[str, str]:
    """URL에서 텍스트 추출. (제목, 본문) 반환"""
    headers = {"User-Agent": "Mozilla/5.0 (compatible; PolicyResearcher/1.0)"}
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
        tag.decompose()
    title = soup.title.string.strip() if soup.title else url
    text = soup.get_text(separator="\n", strip=True)
    # 연속 빈 줄 제거
    lines = [l for l in text.splitlines() if l.strip()]
    return title, "\n".join(lines)


def get_policy_list() -> list:
    if not POLICIES_DIR.exists():
        return []
    return sorted([d.name for d in POLICIES_DIR.iterdir() if d.is_dir()])


# ── 사이드바 ──────────────────────────────────────────────────
policies = get_policy_list()
if not policies:
    st.warning("policies/ 폴더에 정책 폴더가 없습니다.")
    st.stop()

with st.sidebar:
    st.header("📂 정책 선택")
    selected_policy = st.radio(
        "분석할 정책",
        policies,
        format_func=lambda x: x.replace("-", " "),
    )
    st.divider()
    docs = load_policy_documents(selected_policy)
    st.markdown(f"**PDF 문서:** {len(docs)}개")
    with st.expander("PDF 목록 보기"):
        for fname in docs:
            st.text(f"• {fname}")

    # ── 웹 소스 추가 ──────────────────────────────────────────
    st.divider()
    st.markdown("**🌐 웹 소스 추가**")
    st.caption("URL을 입력하면 해당 페이지를 소스로 추가합니다")

    if "web_sources" not in st.session_state:
        st.session_state.web_sources = {}  # {url: {"title": ..., "text": ...}}

    url_input = st.text_input("URL 입력", placeholder="https://...", label_visibility="collapsed")
    if st.button("➕ 소스 추가", use_container_width=True):
        if url_input and url_input.startswith("http"):
            if url_input in st.session_state.web_sources:
                st.warning("이미 추가된 URL입니다.")
            else:
                with st.spinner("페이지를 읽는 중..."):
                    try:
                        title, text = fetch_url_text(url_input)
                        st.session_state.web_sources[url_input] = {"title": title, "text": text}
                        st.success(f"추가됨: {title[:30]}...")
                        st.rerun()
                    except Exception as e:
                        st.error(f"불러오기 실패: {e}")
        else:
            st.warning("올바른 URL을 입력해주세요.")

    if st.session_state.web_sources:
        st.markdown(f"**웹 소스:** {len(st.session_state.web_sources)}개")
        for url, info in list(st.session_state.web_sources.items()):
            col1, col2 = st.columns([4, 1])
            col1.caption(f"🔗 {info['title'][:25]}...")
            if col2.button("✕", key=f"del_{url}", help="삭제"):
                del st.session_state.web_sources[url]
                st.rerun()

    st.divider()
    if st.button("대화 초기화", use_container_width=True):
        st.session_state.messages = []
        st.rerun()

# 세션 초기화
if "messages" not in st.session_state:
    st.session_state.messages = []
if "current_policy" not in st.session_state:
    st.session_state.current_policy = selected_policy
if st.session_state.current_policy != selected_policy:
    st.session_state.messages = []
    st.session_state.current_policy = selected_policy

# ── 헤더 ──────────────────────────────────────────────────────
st.title("📚 정책 리서처")
st.caption(f"현재 정책: **{selected_policy.replace('-', ' ')}** · 문서 {len(docs)}개 분석 중")

# ── 탭 ───────────────────────────────────────────────────────
tab_chat, tab_summary, tab_infographic = st.tabs(["💬 질문하기", "📋 요약 카드", "📊 인포그래픽"])

# ── 탭1: 챗봇 ─────────────────────────────────────────────────
with tab_chat:
    if not st.session_state.messages:
        st.markdown("#### 이런 질문을 해보세요")
        examples = [
            "이 정책에서 마을활동가의 인정 방식은 어떻게 분류되나요?",
            "서울, 경기, 광주의 접근 방식을 비교해주세요",
            "활동가 역량 기준이 어떻게 정의되어 있나요?",
            "기회소득과 인정체계의 연계 방안은 무엇인가요?",
        ]
        cols = st.columns(2)
        for i, q in enumerate(examples):
            if cols[i % 2].button(q, key=f"ex_{i}", use_container_width=True):
                st.session_state.messages.append({"role": "user", "content": q})
                st.rerun()

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("질문을 입력하세요..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("문서를 검색하는 중..."):
                # PDF + 웹소스 합치기
                all_docs = dict(docs)
                for url, info in st.session_state.web_sources.items():
                    all_docs[f"[웹] {info['title']}"] = info["text"]
                context = get_relevant_chunks(prompt, all_docs)
                answer = ask_gemini(prompt, context, st.session_state.messages)
            st.markdown(answer)
        st.session_state.messages.append({"role": "assistant", "content": answer})

# ── 탭2: 요약 카드 ────────────────────────────────────────────
with tab_summary:
    st.markdown("### 📋 문서 핵심 요약")
    st.caption("AI가 각 지역/기관별 핵심 내용을 자동으로 요약합니다.")

    if st.button("🔄 요약 생성하기", type="primary", key="gen_summary"):
        generate_summary_cards.clear()
        st.rerun()

    with st.spinner("문서를 분석하는 중... (처음 1회만 시간이 걸립니다)"):
        combined = "\n\n".join(f"[{fn}]\n{tx[:3000]}" for fn, tx in docs.items())
        cards = generate_summary_cards(selected_policy, combined)

    if cards:
        cols = st.columns(2)
        for i, card in enumerate(cards):
            with cols[i % 2]:
                points_html = "".join(f"<li>{p}</li>" for p in card.get("points", []))
                keyword = card.get("keyword", "")
                st.markdown(f"""
<div class="summary-card">
  <h4>{'🏷️ ' + keyword + ' &nbsp;·&nbsp; ' if keyword else ''}{card.get('title','')}</h4>
  <ul>{points_html}</ul>
</div>
""", unsafe_allow_html=True)
    else:
        st.info("위의 '요약 생성하기' 버튼을 눌러주세요.")

# ── 탭3: 인포그래픽 ───────────────────────────────────────────
with tab_infographic:
    st.markdown("### 📊 정책 현황 인포그래픽")
    st.caption("문서에서 추출한 수치와 비교 데이터를 시각화합니다.")

    if st.button("🔄 인포그래픽 생성하기", type="primary", key="gen_infographic"):
        generate_infographic_data.clear()
        st.rerun()

    with st.spinner("데이터를 분석하는 중..."):
        combined = "\n\n".join(f"[{fn}]\n{tx[:3000]}" for fn, tx in docs.items())
        info = generate_infographic_data(selected_policy, combined)

    if info:
        metrics = info.get("metrics", [])
        if metrics:
            st.markdown("#### 📌 주요 수치")
            mcols = st.columns(min(len(metrics), 3))
            for i, m in enumerate(metrics[:6]):
                with mcols[i % 3]:
                    st.markdown(f"""
<div class="metric-box">
  <div class="num">{m.get('value','–')}<span style="font-size:1rem;color:#555"> {m.get('unit','')}</span></div>
  <div class="label">{m.get('label','')}</div>
</div>
""", unsafe_allow_html=True)

        regions = info.get("regions", [])
        if regions:
            st.markdown("#### 🗺️ 지역별 제도화 수준")
            for r in regions:
                score = int(r.get("score", 50))
                st.markdown(f"""
<div class="region-bar">
  <div class="region-name">📍 {r.get('name','')}</div>
  <div style="font-size:0.83rem;color:#555;margin-bottom:6px">{r.get('approach','')}</div>
  <div class="bar-wrap"><div class="bar-fill" style="width:{score}%"></div></div>
  <div style="text-align:right;font-size:0.78rem;color:#2E7D32;margin-top:2px">{score}점</div>
</div>
""", unsafe_allow_html=True)

        timeline = info.get("timeline", [])
        if timeline:
            st.markdown("#### 📅 주요 연혁")
            for item in timeline:
                c1, c2 = st.columns([1, 5])
                c1.markdown(f"**{item.get('year','')}**")
                c2.markdown(item.get("event", ""))

        issues = info.get("key_issues", [])
        if issues:
            st.markdown("#### ⚡ 핵심 과제")
            icols = st.columns(2)
            for i, issue in enumerate(issues):
                icols[i % 2].markdown(f"- {issue}")
    else:
        st.info("위의 '인포그래픽 생성하기' 버튼을 눌러주세요.")
