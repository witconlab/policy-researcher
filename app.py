import streamlit as st
import google.generativeai as genai
import os, json, io, re
import pdfplumber, requests
from bs4 import BeautifulSoup
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from duckduckgo_search import DDGS

try:
    from youtube_transcript_api import YouTubeTranscriptApi
    YOUTUBE_OK = True
except Exception:
    YOUTUBE_OK = False

st.set_page_config(
    page_title="전남광주 통합특별시 시민주권 정책 공론장",
    page_icon="🏛️",
    layout="wide",
)

st.markdown("""
<style>
/* 전체 */
[data-testid="stSidebar"] { min-width: 220px !important; }

/* 소스 카드 */
.src-card {
    background: #fff;
    border: 1px solid #e0e0e0;
    border-radius: 10px;
    padding: 10px 14px;
    margin-bottom: 6px;
    transition: border-color 0.2s;
}
.src-card.active { border-left: 4px solid #2E7D32; background: #f8fdf8; }
.src-card .src-title { font-size: 0.85rem; font-weight: 600; color: #1B5E20; }
.src-card .src-meta  { font-size: 0.75rem; color: #777; margin-top: 2px; }

/* 채팅 */
.chat-active-sources {
    background: #E8F5E9;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 0.8rem;
    color: #2E7D32;
    margin-bottom: 12px;
}

/* 플래시카드 */
.flashcard {
    background: linear-gradient(135deg,#E8F5E9,#F1F8E9);
    border: 2px solid #2E7D32;
    border-radius: 16px;
    padding: 36px 24px;
    text-align: center;
    min-height: 180px;
    font-size: 1.05rem;
    color: #1B5E20;
    line-height: 1.7;
    margin-bottom: 12px;
}

/* 수치 카드 */
.metric-box {
    background:#E8F5E9; border-radius:10px; padding:18px;
    text-align:center; margin-bottom:8px;
}
.metric-box .num { font-size:2rem; font-weight:700; color:#2E7D32; }
.metric-box .label { font-size:0.82rem; color:#555; margin-top:4px; }

/* 지역 바 */
.region-bar { background:#fff; border:1px solid #C8E6C9; border-radius:8px; padding:14px 18px; margin-bottom:8px; }
.region-bar .rn { font-weight:600; color:#1B5E20; margin-bottom:6px; font-size:0.9rem; }
.bw { background:#E8F5E9; border-radius:4px; height:10px; }
.bf { background:#2E7D32; border-radius:4px; height:10px; }
</style>
""", unsafe_allow_html=True)

# ── API ───────────────────────────────────────────────────────
GOOGLE_API_KEY = st.secrets.get("GOOGLE_API_KEY", os.environ.get("GOOGLE_API_KEY", ""))
if not GOOGLE_API_KEY:
    st.error("Google API 키가 설정되지 않았습니다.")
    st.stop()

genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")
POLICIES_DIR = Path("policies")


# ── 유틸 ──────────────────────────────────────────────────────
def call_gemini(prompt: str) -> str:
    return model.generate_content(prompt).text

def extract_youtube_id(url: str):
    m = re.search(r"(?:v=|youtu\.be/)([A-Za-z0-9_-]{11})", url)
    return m.group(1) if m else None

def fetch_youtube(url: str) -> tuple:
    vid = extract_youtube_id(url)
    if not vid:
        raise ValueError("유튜브 URL을 인식할 수 없습니다.")
    title = f"YouTube: {vid}"
    try:
        resp = requests.get(f"https://www.youtube.com/watch?v={vid}", headers={"User-Agent":"Mozilla/5.0"}, timeout=10)
        m = re.search(r'"title":"([^"]+)"', resp.text)
        if m:
            title = m.group(1)
    except Exception:
        pass
    text = ""
    if YOUTUBE_OK:
        try:
            tlist = YouTubeTranscriptApi.list_transcripts(vid)
            try:
                t = tlist.find_transcript(["ko"])
            except Exception:
                t = tlist.find_generated_transcript(["ko","en"])
            text = " ".join(s["text"] for s in t.fetch())
        except Exception:
            text = f"[유튜브 자막 추출 실패 - URL: {url}]"
    return title, text

def fetch_article(url: str) -> tuple:
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script","style","nav","footer","header","aside"]):
        tag.decompose()
    title = soup.title.string.strip() if soup.title else url
    text  = "\n".join(l for l in soup.get_text(separator="\n",strip=True).splitlines() if l.strip())
    return title, text

def detect_and_fetch(url: str) -> tuple:
    if "youtube.com" in url or "youtu.be" in url:
        return fetch_youtube(url)
    return fetch_article(url)

def search_web(keyword: str, max_results=5) -> list:
    try:
        with DDGS() as d:
            return [{"title":r.get("title",""), "url":r.get("href",""), "snippet":r.get("body","")}
                    for r in d.text(keyword + " 정책", max_results=max_results)]
    except Exception:
        return []


# ── 소스 저장/로드 (policies/{policy}/sources.json) ──────────
def sources_path(policy: str) -> Path:
    return POLICIES_DIR / policy / "sources.json"

def load_saved_sources(policy: str) -> list:
    p = sources_path(policy)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []

def save_sources(policy: str, sources: list):
    p = sources_path(policy)
    p.write_text(json.dumps(sources, ensure_ascii=False, indent=2), encoding="utf-8")


# ── PDF 로드 ──────────────────────────────────────────────────
@st.cache_resource(show_spinner="PDF 자료를 읽는 중...")
def load_pdfs(policy: str) -> dict:
    docs = {}
    policy_path = POLICIES_DIR / policy
    for f in sorted(policy_path.glob("*.pdf")) + sorted(policy_path.glob("*.PDF")):
        try:
            parts = []
            with pdfplumber.open(f) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: parts.append(t)
            if parts:
                docs[f.name] = "\n".join(parts)
        except Exception:
            pass
    return docs

def get_relevant_chunks(query: str, sources_text: dict, max_chars=60000) -> str:
    kws = [w for w in query.lower().split() if len(w) > 1]
    scored = []
    for name, text in sources_text.items():
        score = sum(text.lower().count(k) for k in kws)
        scored.append((score, name, text))
    scored.sort(reverse=True)
    parts, total = [], 0
    for _, name, text in scored:
        chunk = f"[출처: {name}]\n{text[:8000]}\n"
        if total + len(chunk) > max_chars: break
        parts.append(chunk); total += len(chunk)
    return "\n---\n".join(parts)

def ask_gemini(query, context, history):
    system = f"""당신은 정책 문서를 분석하는 전문 리서처입니다.
아래 문서들을 바탕으로만 질문에 답하세요.
- 한국어로 답변
- 근거 출처 인용: ([출처: 파일명])
- 문서에 없으면 "자료에서 확인되지 않습니다" 라고 답변
- 비교 질문은 표 사용

=== 참고 소스 ===
{context}"""
    hist = []
    for msg in history[:-1]:
        hist.append({"role":"user" if msg["role"]=="user" else "model","parts":[msg["content"]]})
    return model.start_chat(history=hist).send_message(f"{system}\n\n질문: {query}").text


# ── 스튜디오 ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def gen_summary_cards(policy, combined):
    prompt = f"""'{policy}' 정책 문서에서 지역/기관별 핵심을 요약.
JSON 배열만 반환:
[{{"title":"지역/기관(30자)","points":["핵심1","핵심2","핵심3"],"keyword":"키워드"}}]
최대 8개. 문서:\n{combined[:40000]}"""
    try:
        t = call_gemini(prompt).strip()
        if "```" in t: t = t.split("```")[1]; t = t[4:] if t.startswith("json") else t
        return json.loads(t)
    except: return []

@st.cache_data(show_spinner=False)
def gen_infographic(policy, combined):
    prompt = f"""정책 문서에서 인포그래픽 데이터 추출. JSON만:
{{"metrics":[{{"label":"","value":"","unit":""}}],
"regions":[{{"name":"","approach":"","score":70}}],
"timeline":[{{"year":"","event":""}}],
"key_issues":[""]}}
문서:\n{combined[:40000]}"""
    try:
        t = call_gemini(prompt).strip()
        if "```" in t: t = t.split("```")[1]; t = t[4:] if t.startswith("json") else t
        return json.loads(t)
    except: return {}

@st.cache_data(show_spinner=False)
def gen_mindmap(policy, combined):
    prompt = f"""정책 문서 핵심개념 마인드맵. Graphviz DOT만 반환:
digraph {{ graph[rankdir=LR] node[shape=rounded,style=filled]
"중심" [fillcolor="#2E7D32",fontcolor=white]
"중심" -> "개념1" ... }}
문서:\n{combined[:30000]}"""
    try:
        t = call_gemini(prompt).strip()
        if "```" in t: t = t.split("```")[1]; t = t[3:] if t.startswith("dot") else t
        return t
    except: return ""

@st.cache_data(show_spinner=False)
def gen_flashcards(policy, combined):
    prompt = f"""'{policy}' 학습용 플래시카드 10개. JSON만:
[{{"question":"질문","answer":"답변(2~3문장)"}}]
문서:\n{combined[:30000]}"""
    try:
        t = call_gemini(prompt).strip()
        if "```" in t: t = t.split("```")[1]; t = t[4:] if t.startswith("json") else t
        return json.loads(t)
    except: return []

@st.cache_data(show_spinner=False)
def gen_quiz(policy, combined):
    prompt = f"""'{policy}' 4지선다 퀴즈 8문제. JSON만:
[{{"question":"질문","options":["①","②","③","④"],"answer":0,"explanation":"해설"}}]
answer=정답인덱스(0~3). 문서:\n{combined[:30000]}"""
    try:
        t = call_gemini(prompt).strip()
        if "```" in t: t = t.split("```")[1]; t = t[4:] if t.startswith("json") else t
        return json.loads(t)
    except: return []

def gen_report(policy, combined):
    return call_gemini(f"""아래 정책 문서로 종합 보고서 작성. 마크다운:
# {policy} 종합 분석 보고서
## 1. 개요 ## 2. 지역별 현황 ## 3. 주요 쟁점 ## 4. 우수 사례 ## 5. 정책 제언
문서:\n{combined[:50000]}""")

def gen_pptx(policy, report):
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    G = RGBColor(0x2E,0x7D,0x32); L = RGBColor(0xE8,0xF5,0xE9)
    def add_slide(ttl, body, cover=False):
        sl = prs.slides.add_slide(prs.slide_layouts[0 if cover else 1])
        for ph in sl.placeholders: ph.text=""
        bg=sl.background; fg=bg.fill; fg.solid(); fg.fore_color.rgb = G if cover else L
        tb=sl.shapes.add_textbox(Inches(.5),Inches(.3),Inches(12.3),Inches(1.2))
        tf=tb.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=ttl; p.font.size=Pt(28 if cover else 22); p.font.bold=True
        p.font.color.rgb=RGBColor(255,255,255) if cover else G
        if body:
            cb=sl.shapes.add_textbox(Inches(.5),Inches(1.7),Inches(12.3),Inches(5.3))
            ctf=cb.text_frame; ctf.word_wrap=True
            for i,ln in enumerate([l for l in body.split("\n") if l.strip()][:18]):
                p2=ctf.paragraphs[0] if i==0 else ctf.add_paragraph()
                p2.text=ln.strip("# ").strip(); p2.font.size=Pt(13)
                p2.font.color.rgb=RGBColor(0x1A,0x1A,0x1A); p2.space_after=Pt(5)
    add_slide(f"{policy}\n종합 분석 보고서","전남광주 통합특별시 시민주권 정책 공론장",cover=True)
    cur_t,cur_b="",[]
    for ln in report.split("\n"):
        if ln.startswith("## "):
            if cur_t: add_slide(cur_t,"\n".join(cur_b))
            cur_t=ln.replace("## ","").strip(); cur_b=[]
        elif not ln.startswith("# "): cur_b.append(ln)
    if cur_t: add_slide(cur_t,"\n".join(cur_b))
    buf=io.BytesIO(); prs.save(buf); buf.seek(0); return buf.read()


def get_policy_list():
    if not POLICIES_DIR.exists(): return []
    return sorted([d.name for d in POLICIES_DIR.iterdir() if d.is_dir()])


# ══════════════════════════════════════════════════════════════
#  세션 초기화
# ══════════════════════════════════════════════════════════════
policies = get_policy_list()
if not policies:
    st.warning("policies/ 폴더에 정책 폴더가 없습니다."); st.stop()

# 사이드바: 정책 선택
with st.sidebar:
    st.title("🏛️ 정책 공론장")
    selected_policy = st.radio("정책 선택", policies,
                                format_func=lambda x: x.replace("-"," "))

# 정책 바뀌면 초기화
if st.session_state.get("_cur_policy") != selected_policy:
    st.session_state._cur_policy   = selected_policy
    st.session_state.messages      = []
    st.session_state.web_sources   = load_saved_sources(selected_policy)  # [{id,type,title,url,text}]
    st.session_state.search_res    = []
    st.session_state.fc_idx        = 0
    st.session_state.fc_show       = False
    st.session_state.qz_idx        = 0
    st.session_state.qz_score      = 0
    st.session_state.qz_answered   = False
    st.session_state.qz_done       = False
    st.session_state.report_text   = ""
    st.session_state.admin_mode    = False

# PDF 로드
pdfs = load_pdfs(selected_policy)

# 소스 체크 상태 초기화
for fname in pdfs:
    k = f"chk_pdf_{fname}"
    if k not in st.session_state: st.session_state[k] = True
for src in st.session_state.web_sources:
    k = f"chk_web_{src['id']}"
    if k not in st.session_state: st.session_state[k] = True


# ══════════════════════════════════════════════════════════════
#  메인 레이아웃
# ══════════════════════════════════════════════════════════════
st.title("🏛️ 전남광주 통합특별시 시민주권 정책 공론장")

tab_chat, tab_studio = st.tabs(["💬 채팅 & 소스", "🎓 스튜디오"])

# ──────────────────────────────────────────────────────────────
#  탭1: 채팅 + 소스 패널
# ──────────────────────────────────────────────────────────────
with tab_chat:
    src_col, chat_col = st.columns([1, 2.5], gap="large")

    # ── 소스 패널 ────────────────────────────────────────────
    with src_col:
        st.markdown("### 📂 소스")

        total_pdf = len(pdfs)
        total_web = len(st.session_state.web_sources)
        checked_pdf = sum(1 for f in pdfs if st.session_state.get(f"chk_pdf_{f}", True))
        checked_web = sum(1 for s in st.session_state.web_sources if st.session_state.get(f"chk_web_{s['id']}", True))
        st.caption(f"PDF {checked_pdf}/{total_pdf} · 웹 {checked_web}/{total_web} 선택됨")

        # 전체 선택/해제
        c1, c2 = st.columns(2)
        if c1.button("전체 선택", use_container_width=True, key="sel_all"):
            for f in pdfs: st.session_state[f"chk_pdf_{f}"] = True
            for s in st.session_state.web_sources: st.session_state[f"chk_web_{s['id']}"] = True
            st.rerun()
        if c2.button("전체 해제", use_container_width=True, key="desel_all"):
            for f in pdfs: st.session_state[f"chk_pdf_{f}"] = False
            for s in st.session_state.web_sources: st.session_state[f"chk_web_{s['id']}"] = False
            st.rerun()

        st.divider()

        # PDF 소스
        if pdfs:
            st.markdown("**📄 PDF 문서**")
            for fname in pdfs:
                key = f"chk_pdf_{fname}"
                checked = st.checkbox(
                    fname[:35] + ("…" if len(fname) > 35 else ""),
                    value=st.session_state.get(key, True),
                    key=key
                )

        # 웹/유튜브 소스
        if st.session_state.web_sources:
            st.divider()
            st.markdown("**🌐 웹 & 유튜브 소스**")
            for src in st.session_state.web_sources:
                key = f"chk_web_{src['id']}"
                icon = "▶️" if src.get("type") == "youtube" else "🔗"
                cols = st.columns([5, 1])
                checked = cols[0].checkbox(
                    f"{icon} {src['title'][:28]}{'…' if len(src['title'])>28 else ''}",
                    value=st.session_state.get(key, True),
                    key=key
                )
                if cols[1].button("✕", key=f"rm_{src['id']}"):
                    st.session_state.web_sources = [s for s in st.session_state.web_sources if s["id"] != src["id"]]
                    save_sources(selected_policy, st.session_state.web_sources)
                    st.rerun()

        st.divider()

        # ── 소스 추가 (관리자 모드) ──────────────────────────
        with st.expander("➕ 소스 추가"):
            add_tab1, add_tab2 = st.tabs(["URL 직접", "키워드 검색"])

            with add_tab1:
                url_in = st.text_input("URL 입력", placeholder="https:// 또는 YouTube 링크", key="url_direct")
                if st.button("추가", key="btn_add_url", use_container_width=True):
                    if url_in.startswith("http"):
                        with st.spinner("소스 읽는 중..."):
                            try:
                                title, text = detect_and_fetch(url_in)
                                src_type = "youtube" if ("youtube" in url_in or "youtu.be" in url_in) else "article"
                                new_src = {"id": str(abs(hash(url_in)))[:10], "type": src_type,
                                           "title": title, "url": url_in, "text": text[:20000]}
                                st.session_state.web_sources.append(new_src)
                                st.session_state[f"chk_web_{new_src['id']}"] = True
                                save_sources(selected_policy, st.session_state.web_sources)
                                st.success(f"추가: {title[:30]}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"실패: {e}")
                    else:
                        st.warning("http로 시작하는 URL을 입력하세요.")

            with add_tab2:
                kw_in = st.text_input("키워드", placeholder="예: 마을활동가 경력인증", key="kw_search")
                if st.button("검색", key="btn_kw", use_container_width=True):
                    with st.spinner("검색 중..."):
                        st.session_state.search_res = search_web(kw_in)
                for r in st.session_state.get("search_res", []):
                    c1, c2 = st.columns([4, 1])
                    c1.markdown(f"**{r['title'][:30]}**\n\n{r['snippet'][:60]}...")
                    if c2.button("✚", key=f"srch_{r['url'][:20]}"):
                        with st.spinner("추가 중..."):
                            try:
                                title, text = fetch_article(r["url"])
                            except Exception:
                                title, text = r["title"], r["snippet"]
                            new_src = {"id": str(abs(hash(r["url"])))[:10], "type": "article",
                                       "title": title, "url": r["url"], "text": text[:20000]}
                            st.session_state.web_sources.append(new_src)
                            st.session_state[f"chk_web_{new_src['id']}"] = True
                            save_sources(selected_policy, st.session_state.web_sources)
                            st.rerun()

    # ── 채팅 패널 ────────────────────────────────────────────
    with chat_col:
        # 현재 활성 소스 표시
        active_pdfs = [f for f in pdfs if st.session_state.get(f"chk_pdf_{f}", True)]
        active_web  = [s for s in st.session_state.web_sources if st.session_state.get(f"chk_web_{s['id']}", True)]
        active_count = len(active_pdfs) + len(active_web)

        if active_count == 0:
            st.warning("⚠️ 선택된 소스가 없습니다. 왼쪽에서 소스를 선택해주세요.")
        else:
            st.markdown(f'<div class="chat-active-sources">📎 현재 <b>{active_count}개 소스</b> 참고 중 · PDF {len(active_pdfs)}개 · 웹 {len(active_web)}개</div>', unsafe_allow_html=True)

        # 예시 질문
        if not st.session_state.messages:
            st.markdown("#### 이런 질문을 해보세요")
            examples = [
                "마을활동가의 인정 방식은 어떻게 분류되나요?",
                "서울, 경기, 광주의 접근 방식을 비교해주세요",
                "활동가 역량 기준이 어떻게 정의되어 있나요?",
                "기회소득과 인정체계의 연계 방안은 무엇인가요?",
            ]
            cols = st.columns(2)
            for i, q in enumerate(examples):
                if cols[i%2].button(q, key=f"ex_{i}", use_container_width=True):
                    st.session_state.messages.append({"role":"user","content":q})
                    st.rerun()

        # 대화 출력
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # 입력
        if prompt := st.chat_input("선택된 소스에 대해 질문하세요..."):
            if active_count == 0:
                st.warning("소스를 먼저 선택해주세요.")
            else:
                st.session_state.messages.append({"role":"user","content":prompt})
                with st.chat_message("user"):
                    st.markdown(prompt)
                with st.chat_message("assistant"):
                    with st.spinner("소스를 검색하는 중..."):
                        all_docs = {f: pdfs[f] for f in active_pdfs}
                        for s in active_web:
                            all_docs[s["title"]] = s["text"]
                        context = get_relevant_chunks(prompt, all_docs)
                        answer  = ask_gemini(prompt, context, st.session_state.messages)
                    st.markdown(answer)
                st.session_state.messages.append({"role":"assistant","content":answer})

        if st.session_state.messages:
            if st.button("🗑️ 대화 초기화", key="clear_chat"):
                st.session_state.messages = []
                st.rerun()

# ──────────────────────────────────────────────────────────────
#  탭2: 스튜디오
# ──────────────────────────────────────────────────────────────
with tab_studio:
    st.markdown("### 🎓 스튜디오")
    combined = "\n\n".join(f"[{fn}]\n{tx[:3000]}" for fn, tx in pdfs.items())

    st1, st2, st3, st4 = st.tabs(["📋 요약 카드", "📊 인포그래픽", "🗺️ 마인드맵", "🃏 플래시카드 · 퀴즈"])

    # 요약 카드
    with st1:
        if st.button("🔄 요약 생성", type="primary"):
            gen_summary_cards.clear(); st.rerun()
        with st.spinner("분석 중..."):
            cards = gen_summary_cards(selected_policy, combined)
        if cards:
            cols = st.columns(2)
            for i, c in enumerate(cards):
                with cols[i%2]:
                    pts = "".join(f"<li>{p}</li>" for p in c.get("points",[]))
                    kw = c.get("keyword","")
                    st.markdown(f"""<div style="background:#f8fdf8;border-left:4px solid #2E7D32;border-radius:8px;padding:14px 18px;margin-bottom:10px">
<b style="color:#1B5E20">{'🏷️ '+kw+' · ' if kw else ''}{c.get('title','')}</b>
<ul style="margin:8px 0 0 0;padding-left:18px;color:#333;font-size:.87rem;line-height:1.7">{pts}</ul></div>""", unsafe_allow_html=True)
        else:
            st.info("생성 버튼을 눌러주세요.")

    # 인포그래픽
    with st2:
        if st.button("🔄 인포그래픽 생성", type="primary"):
            gen_infographic.clear(); st.rerun()
        with st.spinner("데이터 분석 중..."):
            info = gen_infographic(selected_policy, combined)
        if info:
            mets = info.get("metrics",[])
            if mets:
                st.markdown("#### 📌 주요 수치")
                mc = st.columns(min(len(mets),3))
                for i,m in enumerate(mets[:6]):
                    with mc[i%3]:
                        st.markdown(f'<div class="metric-box"><div class="num">{m.get("value","–")}<span style="font-size:1rem;color:#555"> {m.get("unit","")}</span></div><div class="label">{m.get("label","")}</div></div>', unsafe_allow_html=True)
            regs = info.get("regions",[])
            if regs:
                st.markdown("#### 🗺️ 지역별 제도화 수준")
                for r in regs:
                    sc = int(r.get("score",50))
                    st.markdown(f'<div class="region-bar"><div class="rn">📍 {r.get("name","")}</div><div style="font-size:.83rem;color:#555;margin-bottom:6px">{r.get("approach","")}</div><div class="bw"><div class="bf" style="width:{sc}%"></div></div><div style="text-align:right;font-size:.78rem;color:#2E7D32;margin-top:2px">{sc}점</div></div>', unsafe_allow_html=True)
            tl = info.get("timeline",[])
            if tl:
                st.markdown("#### 📅 주요 연혁")
                for item in tl:
                    a,b = st.columns([1,5])
                    a.markdown(f"**{item.get('year','')}**")
                    b.markdown(item.get("event",""))
            issues = info.get("key_issues",[])
            if issues:
                st.markdown("#### ⚡ 핵심 과제")
                ic = st.columns(2)
                for i,iss in enumerate(issues): ic[i%2].markdown(f"- {iss}")
        else:
            st.info("생성 버튼을 눌러주세요.")

    # 마인드맵
    with st3:
        if st.button("🔄 마인드맵 생성", type="primary"):
            gen_mindmap.clear(); st.rerun()
        with st.spinner("마인드맵 생성 중..."):
            dot = gen_mindmap(selected_policy, combined)
        if dot:
            try: st.graphviz_chart(dot, use_container_width=True)
            except Exception as e: st.error(f"렌더링 오류: {e}"); st.code(dot)
        else:
            st.info("생성 버튼을 눌러주세요.")

    # 플래시카드 + 퀴즈 + 보고서
    with st4:
        lt, rt = st.columns(2)

        # 플래시카드
        with lt:
            st.markdown("#### 🃏 플래시카드")
            if st.button("🔄 카드 생성", type="primary", key="fc_gen"):
                gen_flashcards.clear()
                st.session_state.fc_idx  = 0
                st.session_state.fc_show = False
                st.rerun()
            with st.spinner("생성 중..."): fc = gen_flashcards(selected_policy, combined)
            if fc:
                n = len(fc); idx = st.session_state.fc_idx % n; card = fc[idx]
                st.markdown(f"**{idx+1}/{n}**")
                if not st.session_state.fc_show:
                    st.markdown(f'<div class="flashcard">❓ {card["question"]}</div>', unsafe_allow_html=True)
                    if st.button("답 보기 👁️", use_container_width=True): st.session_state.fc_show=True; st.rerun()
                else:
                    st.markdown(f'<div class="flashcard">💡 {card["answer"]}</div>', unsafe_allow_html=True)
                    c1,c2 = st.columns(2)
                    if c1.button("⬅️ 이전", use_container_width=True): st.session_state.fc_idx=(idx-1)%n; st.session_state.fc_show=False; st.rerun()
                    if c2.button("다음 ➡️", use_container_width=True): st.session_state.fc_idx=(idx+1)%n; st.session_state.fc_show=False; st.rerun()
            else: st.info("생성 버튼을 눌러주세요.")

        # 퀴즈
        with rt:
            st.markdown("#### 🧠 퀴즈")
            if st.button("🔄 퀴즈 생성", type="primary", key="qz_gen"):
                gen_quiz.clear()
                st.session_state.qz_idx=0; st.session_state.qz_score=0
                st.session_state.qz_answered=False; st.session_state.qz_done=False; st.rerun()
            with st.spinner("생성 중..."): qz = gen_quiz(selected_policy, combined)
            if qz and not st.session_state.qz_done:
                qi = st.session_state.qz_idx
                if qi < len(qz):
                    q = qz[qi]
                    st.markdown(f"**{qi+1}/{len(qz)}** | 점수: {st.session_state.qz_score}")
                    st.markdown(f"**{q['question']}**")
                    for oi,opt in enumerate(q["options"]):
                        if st.button(opt, key=f"opt_{qi}_{oi}", use_container_width=True, disabled=st.session_state.qz_answered):
                            st.session_state.qz_answered=True
                            if oi==q["answer"]: st.session_state.qz_score+=1; st.success("✅ 정답!")
                            else: st.error(f"❌ 오답! 정답: {q['options'][q['answer']]}")
                            st.info(f"📖 {q.get('explanation','')}")
                    if st.session_state.qz_answered:
                        if st.button("다음 ➡️", use_container_width=True, key=f"qn_{qi}"):
                            st.session_state.qz_idx+=1; st.session_state.qz_answered=False
                            if st.session_state.qz_idx>=len(qz): st.session_state.qz_done=True
                            st.rerun()
                else: st.session_state.qz_done=True; st.rerun()
            elif st.session_state.qz_done and qz:
                pct=int(st.session_state.qz_score/len(qz)*100)
                st.markdown(f"## 🎉 완료! {st.session_state.qz_score}/{len(qz)} ({pct}점)")
                if pct>=80: st.success("우수!")
                elif pct>=50: st.warning("플래시카드로 복습해보세요.")
                else: st.error("요약 카드부터 다시 읽어보세요.")
                if st.button("다시 도전"):
                    st.session_state.qz_idx=0; st.session_state.qz_score=0
                    st.session_state.qz_answered=False; st.session_state.qz_done=False; st.rerun()
            else: st.info("생성 버튼을 눌러주세요.")

        st.divider()

        # 보고서 + PPT
        st.markdown("#### 📄 보고서 & 슬라이드")
        if st.button("📝 보고서 생성", type="primary"):
            with st.spinner("보고서 작성 중... (30초~1분)"):
                full = "\n\n".join(f"[{fn}]\n{tx[:5000]}" for fn,tx in pdfs.items())
                st.session_state.report_text = gen_report(selected_policy, full)
        if st.session_state.report_text:
            st.markdown(st.session_state.report_text)
            st.divider()
            c1,c2 = st.columns(2)
            c1.download_button("📥 보고서 (.md)", data=st.session_state.report_text,
                file_name=f"{selected_policy}_보고서.md", mime="text/markdown", use_container_width=True)
            if c2.button("📊 PPT 생성", use_container_width=True):
                with st.spinner("PPT 생성 중..."):
                    pptx_b = gen_pptx(selected_policy, st.session_state.report_text)
                st.download_button("⬇️ PPT 다운로드", data=pptx_b,
                    file_name=f"{selected_policy}_발표자료.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True)
        else:
            st.info("'보고서 생성' 버튼을 눌러주세요.")
