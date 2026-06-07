import streamlit as st
import google.generativeai as genai
import os
import json
import pdfplumber
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="전남광주 통합특별시 시민주권 정책 공론장",
    page_icon="🏛️",
    layout="wide",
)

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
.flashcard {
    background: linear-gradient(135deg, #E8F5E9, #F1F8E9);
    border: 2px solid #2E7D32;
    border-radius: 16px;
    padding: 40px 30px;
    text-align: center;
    min-height: 200px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 16px;
    font-size: 1.1rem;
    color: #1B5E20;
    line-height: 1.7;
}
.quiz-option {
    background: #f8fdf8;
    border: 1px solid #C8E6C9;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 8px;
    cursor: pointer;
}
.search-result {
    background: #fff;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 8px;
}
.search-result .title { font-weight: 600; color: #1B5E20; font-size: 0.9rem; }
.search-result .snippet { font-size: 0.83rem; color: #555; margin-top: 4px; }
.search-result .url { font-size: 0.75rem; color: #2E7D32; margin-top: 4px; }
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
def load_policy_documents(policy_name: str) -> dict:
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


def get_relevant_chunks(query: str, docs: dict, max_chars: int = 60000) -> str:
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


def call_gemini(prompt: str) -> str:
    return model.generate_content(prompt).text


# ── 웹 소스 ───────────────────────────────────────────────────
def fetch_url_text(url: str) -> tuple:
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
        tag.decompose()
    title = soup.title.string.strip() if soup.title else url
    text = "\n".join(l for l in soup.get_text(separator="\n", strip=True).splitlines() if l.strip())
    return title, text


def search_web(keyword: str, max_results: int = 5) -> list:
    """DuckDuckGo로 키워드 검색 후 결과 반환"""
    results = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(keyword + " 정책 마을 한국", max_results=max_results):
                results.append({
                    "title": r.get("title", ""),
                    "url": r.get("href", ""),
                    "snippet": r.get("body", "")
                })
    except Exception:
        pass
    return results


# ── 스튜디오 기능들 ───────────────────────────────────────────
@st.cache_data(show_spinner=False)
def generate_summary_cards(policy_name: str, doc_texts: str) -> list:
    prompt = f"""아래는 '{policy_name}' 관련 정책 문서들입니다.
지역 또는 연구 기관별로 핵심 내용을 요약해주세요.
반드시 아래 JSON 배열 형식으로만 응답하세요:
[{{"title": "지역/기관명 (30자 이내)", "points": ["핵심1", "핵심2", "핵심3"], "keyword": "키워드"}}]
최대 8개 카드.
문서:\n{doc_texts[:40000]}"""
    try:
        text = call_gemini(prompt).strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def generate_infographic_data(policy_name: str, doc_texts: str) -> dict:
    prompt = f"""아래 정책 문서에서 인포그래픽 데이터를 추출하세요.
JSON 형식으로만 응답:
{{"metrics":[{{"label":"","value":"","unit":"","source":""}}],
"regions":[{{"name":"","approach":"","score":75}}],
"timeline":[{{"year":"","event":""}}],
"key_issues":[""]}}
metrics 최대 6개, regions 최대 6개, timeline 최대 6개.
문서:\n{doc_texts[:40000]}"""
    try:
        text = call_gemini(prompt).strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return {}


@st.cache_data(show_spinner=False)
def generate_mindmap(policy_name: str, doc_texts: str) -> str:
    prompt = f"""아래 정책 문서에서 핵심 개념 관계도(마인드맵)를 만들어주세요.
반드시 Graphviz DOT 언어 형식으로만 응답하세요. 다른 텍스트 없이 DOT 코드만:

digraph mindmap {{
    graph [rankdir=LR, fontname="NanumGothic"]
    node [shape=rounded, style=filled, fontname="NanumGothic"]
    "중심주제" [fillcolor="#2E7D32", fontcolor=white, fontsize=14]
    "중심주제" -> "하위개념1"
    ...
}}

중심 노드 1개, 1단계 노드 4~6개, 2단계 노드 각 2~3개.
문서:\n{doc_texts[:30000]}"""
    try:
        text = call_gemini(prompt).strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("dot"):
                text = text[3:]
        return text
    except Exception:
        return ""


@st.cache_data(show_spinner=False)
def generate_flashcards(policy_name: str, doc_texts: str) -> list:
    prompt = f"""아래 정책 문서에서 학습용 플래시카드를 만들어주세요.
JSON 배열 형식으로만 응답:
[{{"question": "질문", "answer": "답변 (2~3문장)"}}]
10개 생성. 핵심 개념, 지역별 특징, 정책 용어 중심으로.
문서:\n{doc_texts[:30000]}"""
    try:
        text = call_gemini(prompt).strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def generate_quiz(policy_name: str, doc_texts: str) -> list:
    prompt = f"""아래 정책 문서에서 4지선다 퀴즈를 만들어주세요.
JSON 배열 형식으로만 응답:
[{{"question":"질문","options":["①답1","②답2","③답3","④답4"],"answer":0,"explanation":"해설"}}]
answer는 정답 인덱스(0~3). 8문제 생성.
문서:\n{doc_texts[:30000]}"""
    try:
        text = call_gemini(prompt).strip()
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception:
        return []


def generate_report(policy_name: str, doc_texts: str) -> str:
    prompt = f"""아래 정책 문서를 바탕으로 종합 보고서를 작성해주세요.
마크다운 형식으로 작성:
# {policy_name} 종합 분석 보고서

## 1. 개요
## 2. 지역별 현황 비교
## 3. 주요 쟁점과 과제
## 4. 우수 사례
## 5. 정책 제언

각 섹션은 구체적인 내용으로 작성. 문서에 없는 내용은 포함하지 마세요.
문서:\n{doc_texts[:50000]}"""
    return call_gemini(prompt)


def generate_pptx(policy_name: str, report_text: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    green = RGBColor(0x2E, 0x7D, 0x32)
    light = RGBColor(0xE8, 0xF5, 0xE9)

    def add_slide(title_text, content_text, is_title_slide=False):
        layout = prs.slide_layouts[0] if is_title_slide else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        for ph in slide.placeholders:
            ph.text = ""
        # 배경
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = light if not is_title_slide else green

        # 제목
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1.2))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(28 if is_title_slide else 22)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF,0xFF,0xFF) if is_title_slide else green

        if content_text:
            cBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.3))
            ctf = cBox.text_frame
            ctf.word_wrap = True
            for i, line in enumerate(content_text.split("\n")[:20]):
                if not line.strip():
                    continue
                p2 = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
                p2.text = line.strip("# ").strip()
                p2.font.size = Pt(14)
                p2.font.color.rgb = RGBColor(0x1A,0x1A,0x1A)
                p2.space_after = Pt(6)

    # 표지
    add_slide(f"{policy_name}\n종합 분석 보고서", "전남광주 통합특별시 시민주권 정책 공론장", is_title_slide=True)

    # 섹션별 슬라이드
    current_title = ""
    current_content = []
    for line in report_text.split("\n"):
        if line.startswith("## "):
            if current_title:
                add_slide(current_title, "\n".join(current_content))
            current_title = line.replace("## ", "").strip()
            current_content = []
        elif line.startswith("# "):
            pass
        else:
            current_content.append(line)
    if current_title:
        add_slide(current_title, "\n".join(current_content))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


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

    # ── 웹 소스 (키워드 검색) ─────────────────────────────────
    st.divider()
    st.markdown("**🌐 웹 소스 추가**")
    st.caption("키워드로 검색해서 소스를 추가합니다")

    if "web_sources" not in st.session_state:
        st.session_state.web_sources = {}
    if "search_results" not in st.session_state:
        st.session_state.search_results = []

    keyword_input = st.text_input("키워드 검색", placeholder="예: 마을활동가 경력인증", label_visibility="collapsed")

    if st.button("🔍 검색", use_container_width=True):
        if keyword_input.strip():
            with st.spinner("검색 중..."):
                st.session_state.search_results = search_web(keyword_input.strip())
        else:
            st.warning("키워드를 입력해주세요.")

    if st.session_state.search_results:
        st.markdown("**검색 결과** (✚ 클릭해서 추가)")
        for i, r in enumerate(st.session_state.search_results):
            col1, col2 = st.columns([5, 1])
            col1.markdown(f"<div class='search-result'><div class='title'>{r['title'][:40]}</div><div class='snippet'>{r['snippet'][:80]}...</div><div class='url'>{r['url'][:50]}</div></div>", unsafe_allow_html=True)
            if col2.button("✚", key=f"add_{i}"):
                if r["url"] not in st.session_state.web_sources:
                    with st.spinner("페이지 읽는 중..."):
                        try:
                            title, text = fetch_url_text(r["url"])
                            st.session_state.web_sources[r["url"]] = {"title": title, "text": text}
                            st.success(f"추가됨!")
                            st.rerun()
                        except Exception:
                            st.session_state.web_sources[r["url"]] = {"title": r["title"], "text": r["snippet"]}
                            st.rerun()

    if st.session_state.web_sources:
        st.markdown(f"**추가된 웹 소스:** {len(st.session_state.web_sources)}개")
        for url, info in list(st.session_state.web_sources.items()):
            c1, c2 = st.columns([4, 1])
            c1.caption(f"🔗 {info['title'][:25]}...")
            if c2.button("✕", key=f"del_{url}"):
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
if "fc_index" not in st.session_state:
    st.session_state.fc_index = 0
if "fc_show_answer" not in st.session_state:
    st.session_state.fc_show_answer = False
if "quiz_index" not in st.session_state:
    st.session_state.quiz_index = 0
if "quiz_score" not in st.session_state:
    st.session_state.quiz_score = 0
if "quiz_answered" not in st.session_state:
    st.session_state.quiz_answered = False
if "quiz_done" not in st.session_state:
    st.session_state.quiz_done = False

# ── 헤더 ──────────────────────────────────────────────────────
st.title("🏛️ 전남광주 통합특별시 시민주권 정책 공론장")
st.caption(f"현재 정책: **{selected_policy.replace('-', ' ')}** · 문서 {len(docs)}개 분석 중")

combined = "\n\n".join(f"[{fn}]\n{tx[:3000]}" for fn, tx in docs.items())

# ── 탭 ───────────────────────────────────────────────────────
tab_chat, tab_summary, tab_infographic, tab_studio = st.tabs([
    "💬 질문하기", "📋 요약 카드", "📊 인포그래픽", "🎓 스튜디오"
])

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
    with st.spinner("분석 중..."):
        cards = generate_summary_cards(selected_policy, combined)
    if cards:
        cols = st.columns(2)
        for i, card in enumerate(cards):
            with cols[i % 2]:
                points_html = "".join(f"<li>{p}</li>" for p in card.get("points", []))
                kw = card.get("keyword", "")
                st.markdown(f"""<div class="summary-card"><h4>{'🏷️ ' + kw + ' &nbsp;·&nbsp; ' if kw else ''}{card.get('title','')}</h4><ul>{points_html}</ul></div>""", unsafe_allow_html=True)
    else:
        st.info("위의 '요약 생성하기' 버튼을 눌러주세요.")

# ── 탭3: 인포그래픽 ───────────────────────────────────────────
with tab_infographic:
    st.markdown("### 📊 정책 현황 인포그래픽")
    if st.button("🔄 인포그래픽 생성하기", type="primary", key="gen_infographic"):
        generate_infographic_data.clear()
        st.rerun()
    with st.spinner("데이터 분석 중..."):
        info = generate_infographic_data(selected_policy, combined)
    if info:
        metrics = info.get("metrics", [])
        if metrics:
            st.markdown("#### 📌 주요 수치")
            mcols = st.columns(min(len(metrics), 3))
            for i, m in enumerate(metrics[:6]):
                with mcols[i % 3]:
                    st.markdown(f"""<div class="metric-box"><div class="num">{m.get('value','–')}<span style="font-size:1rem;color:#555"> {m.get('unit','')}</span></div><div class="label">{m.get('label','')}</div></div>""", unsafe_allow_html=True)
        regions = info.get("regions", [])
        if regions:
            st.markdown("#### 🗺️ 지역별 제도화 수준")
            for r in regions:
                score = int(r.get("score", 50))
                st.markdown(f"""<div class="region-bar"><div class="region-name">📍 {r.get('name','')}</div><div style="font-size:0.83rem;color:#555;margin-bottom:6px">{r.get('approach','')}</div><div class="bar-wrap"><div class="bar-fill" style="width:{score}%"></div></div><div style="text-align:right;font-size:0.78rem;color:#2E7D32;margin-top:2px">{score}점</div></div>""", unsafe_allow_html=True)
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

# ── 탭4: 스튜디오 ─────────────────────────────────────────────
with tab_studio:
    st.markdown("### 🎓 스튜디오")
    st.caption("정책 문서를 다양한 형식으로 변환하고 학습합니다.")

    studio_tab1, studio_tab2, studio_tab3 = st.tabs(["🗺️ 마인드맵", "🃏 플래시카드 · 퀴즈", "📄 보고서 · 슬라이드"])

    # ── 마인드맵 ────────────────────────────────────────────
    with studio_tab1:
        st.markdown("#### 🗺️ 개념 관계도 (마인드맵)")
        st.caption("문서의 핵심 개념과 관계를 시각화합니다.")
        if st.button("🔄 마인드맵 생성", type="primary", key="gen_mindmap"):
            generate_mindmap.clear()
            st.rerun()
        with st.spinner("마인드맵을 생성하는 중..."):
            dot_code = generate_mindmap(selected_policy, combined)
        if dot_code:
            try:
                st.graphviz_chart(dot_code, use_container_width=True)
            except Exception as e:
                st.error(f"마인드맵 렌더링 오류: {e}")
                st.code(dot_code, language="dot")
        else:
            st.info("'마인드맵 생성' 버튼을 눌러주세요.")

    # ── 플래시카드 + 퀴즈 ───────────────────────────────────
    with studio_tab2:
        fc_col, quiz_col = st.columns(2)

        with fc_col:
            st.markdown("#### 🃏 플래시카드")
            if st.button("🔄 카드 생성", type="primary", key="gen_fc"):
                generate_flashcards.clear()
                st.session_state.fc_index = 0
                st.session_state.fc_show_answer = False
                st.rerun()
            with st.spinner("플래시카드 생성 중..."):
                flashcards = generate_flashcards(selected_policy, combined)
            if flashcards:
                total = len(flashcards)
                idx = st.session_state.fc_index % total
                card = flashcards[idx]
                st.markdown(f"**{idx+1} / {total}**")
                if not st.session_state.fc_show_answer:
                    st.markdown(f'<div class="flashcard">❓ {card["question"]}</div>', unsafe_allow_html=True)
                    if st.button("답 보기 👁️", use_container_width=True, key="fc_reveal"):
                        st.session_state.fc_show_answer = True
                        st.rerun()
                else:
                    st.markdown(f'<div class="flashcard">💡 {card["answer"]}</div>', unsafe_allow_html=True)
                    c1, c2 = st.columns(2)
                    if c1.button("⬅️ 이전", use_container_width=True, key="fc_prev"):
                        st.session_state.fc_index = (idx - 1) % total
                        st.session_state.fc_show_answer = False
                        st.rerun()
                    if c2.button("다음 ➡️", use_container_width=True, key="fc_next"):
                        st.session_state.fc_index = (idx + 1) % total
                        st.session_state.fc_show_answer = False
                        st.rerun()
            else:
                st.info("'카드 생성' 버튼을 눌러주세요.")

        with quiz_col:
            st.markdown("#### 🧠 퀴즈")
            if st.button("🔄 퀴즈 생성", type="primary", key="gen_quiz"):
                generate_quiz.clear()
                st.session_state.quiz_index = 0
                st.session_state.quiz_score = 0
                st.session_state.quiz_answered = False
                st.session_state.quiz_done = False
                st.rerun()
            with st.spinner("퀴즈 생성 중..."):
                quiz_list = generate_quiz(selected_policy, combined)
            if quiz_list and not st.session_state.quiz_done:
                total_q = len(quiz_list)
                qi = st.session_state.quiz_index
                if qi < total_q:
                    q = quiz_list[qi]
                    st.markdown(f"**문제 {qi+1} / {total_q}** &nbsp;|&nbsp; 점수: {st.session_state.quiz_score}점")
                    st.markdown(f"**{q['question']}**")
                    for oi, opt in enumerate(q["options"]):
                        if st.button(opt, key=f"opt_{qi}_{oi}", use_container_width=True, disabled=st.session_state.quiz_answered):
                            st.session_state.quiz_answered = True
                            if oi == q["answer"]:
                                st.session_state.quiz_score += 1
                                st.success("✅ 정답!")
                            else:
                                st.error(f"❌ 오답! 정답: {q['options'][q['answer']]}")
                            st.info(f"📖 {q.get('explanation','')}")
                    if st.session_state.quiz_answered:
                        if st.button("다음 문제 ➡️", use_container_width=True, key=f"quiz_next_{qi}"):
                            st.session_state.quiz_index += 1
                            st.session_state.quiz_answered = False
                            if st.session_state.quiz_index >= total_q:
                                st.session_state.quiz_done = True
                            st.rerun()
                else:
                    st.session_state.quiz_done = True
                    st.rerun()
            elif st.session_state.quiz_done and quiz_list:
                total_q = len(quiz_list)
                score = st.session_state.quiz_score
                pct = int(score / total_q * 100)
                st.markdown(f"## 🎉 퀴즈 완료!")
                st.markdown(f"**{total_q}문제 중 {score}문제 정답 ({pct}점)**")
                if pct >= 80:
                    st.success("우수! 정책을 잘 이해하고 있습니다.")
                elif pct >= 50:
                    st.warning("보통. 플래시카드로 복습해보세요.")
                else:
                    st.error("요약 카드부터 다시 읽어보세요.")
                if st.button("다시 도전", use_container_width=True):
                    st.session_state.quiz_index = 0
                    st.session_state.quiz_score = 0
                    st.session_state.quiz_answered = False
                    st.session_state.quiz_done = False
                    st.rerun()
            else:
                st.info("'퀴즈 생성' 버튼을 눌러주세요.")

    # ── 보고서 + 슬라이드 ────────────────────────────────────
    with studio_tab3:
        st.markdown("#### 📄 보고서 및 슬라이드 생성")
        st.caption("문서를 분석해 종합 보고서와 PPT를 자동 생성합니다.")

        if "report_text" not in st.session_state:
            st.session_state.report_text = ""

        if st.button("📝 보고서 생성", type="primary", key="gen_report"):
            with st.spinner("보고서를 작성하는 중... (30초~1분 소요)"):
                full_combined = "\n\n".join(f"[{fn}]\n{tx[:5000]}" for fn, tx in docs.items())
                st.session_state.report_text = generate_report(selected_policy, full_combined)

        if st.session_state.report_text:
            st.markdown(st.session_state.report_text)
            st.divider()
            col1, col2 = st.columns(2)
            col1.download_button(
                "📥 보고서 다운로드 (.md)",
                data=st.session_state.report_text,
                file_name=f"{selected_policy}_보고서.md",
                mime="text/markdown",
                use_container_width=True,
            )
            if col2.button("📊 PPT 생성 및 다운로드", use_container_width=True):
                with st.spinner("PPT 생성 중..."):
                    pptx_bytes = generate_pptx(selected_policy, st.session_state.report_text)
                st.download_button(
                    "⬇️ PPT 다운로드",
                    data=pptx_bytes,
                    file_name=f"{selected_policy}_발표자료.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
        else:
            st.info("'보고서 생성' 버튼을 눌러주세요.")
