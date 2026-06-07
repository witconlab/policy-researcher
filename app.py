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

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="전남광주 통합특별시 시민주권 정책 공론장",
    page_icon="🏛️",
    layout="wide",
)

st.markdown("""
<style>
/* ── 전체 배경 ── */
.stApp { background: #F2F3F5; }
[data-testid="stSidebar"] { background: #FFFFFF; min-width: 260px !important; }

/* ── 카카오톡 채팅창 ── */
.chat-wrap {
    background: #B2C7D9;
    border-radius: 16px;
    padding: 20px 16px;
    min-height: 520px;
    max-height: 600px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 14px;
}

/* 사용자 버블 (오른쪽, 노란색) */
.bubble-user {
    display: flex;
    justify-content: flex-end;
    align-items: flex-end;
    gap: 6px;
}
.bubble-user .bubble {
    background: #FEE500;
    color: #1A1A1A;
    border-radius: 18px 18px 4px 18px;
    padding: 12px 16px;
    max-width: 68%;
    font-size: 1.05rem;
    line-height: 1.6;
    box-shadow: 0 1px 3px rgba(0,0,0,.12);
    word-break: break-word;
}
.bubble-user .time {
    font-size: 0.72rem;
    color: #555;
    margin-bottom: 4px;
    white-space: nowrap;
}

/* AI 버블 (왼쪽, 흰색) */
.bubble-ai {
    display: flex;
    justify-content: flex-start;
    align-items: flex-start;
    gap: 8px;
}
.bubble-ai .avatar {
    width: 40px; height: 40px;
    border-radius: 12px;
    background: #2E7D32;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.2rem;
    flex-shrink: 0;
    margin-top: 2px;
}
.bubble-ai .bubble-body { display: flex; flex-direction: column; gap: 2px; max-width: 72%; }
.bubble-ai .sender { font-size: 0.78rem; font-weight: 600; color: #333; margin-left: 2px; }
.bubble-ai .bubble {
    background: #FFFFFF;
    color: #1A1A1A;
    border-radius: 4px 18px 18px 18px;
    padding: 13px 16px;
    font-size: 1.05rem;
    line-height: 1.7;
    box-shadow: 0 1px 3px rgba(0,0,0,.10);
    word-break: break-word;
}
.bubble-ai .time {
    font-size: 0.72rem;
    color: #666;
    margin-top: 2px;
    margin-left: 2px;
}

/* ── 소스 카드 ── */
.src-header { font-size:.78rem; font-weight:700; color:#888; letter-spacing:.04em; margin:10px 0 4px; }
.src-item {
    display:flex; align-items:center; gap:8px;
    background:#fff; border:1px solid #e8e8e8;
    border-radius:10px; padding:9px 12px; margin-bottom:5px;
    font-size:.85rem; color:#222;
}
.src-item.active { border-left:4px solid #2E7D32; background:#f8fdf8; }
.src-icon { font-size:1rem; flex-shrink:0; }

/* ── 입력창 ── */
.stChatInput textarea { font-size:1rem !important; }

/* ── 관리자 패널 ── */
.admin-badge {
    background:#1B5E20; color:#fff; border-radius:6px;
    padding:3px 10px; font-size:.78rem; font-weight:700;
    display:inline-block; margin-bottom:10px;
}

/* ── 메트릭/인포 ── */
.metric-box { background:#E8F5E9;border-radius:10px;padding:16px;text-align:center;margin-bottom:8px; }
.metric-box .num { font-size:1.8rem;font-weight:700;color:#2E7D32; }
.metric-box .lbl { font-size:.8rem;color:#555;margin-top:3px; }
.flashcard { background:linear-gradient(135deg,#E8F5E9,#F1F8E9);border:2px solid #2E7D32;border-radius:16px;padding:36px 24px;text-align:center;min-height:160px;font-size:1.05rem;color:#1B5E20;line-height:1.7;margin-bottom:12px; }
</style>
""", unsafe_allow_html=True)

# ── API ───────────────────────────────────────────────────────
GOOGLE_API_KEY = st.secrets.get("GOOGLE_API_KEY", os.environ.get("GOOGLE_API_KEY",""))
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin1234")

if not GOOGLE_API_KEY:
    st.error("Google API 키가 없습니다."); st.stop()

genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")
POLICIES_DIR = Path("policies")


# ── 유틸 ──────────────────────────────────────────────────────
def call_gemini(prompt):
    return model.generate_content(prompt).text

def extract_youtube_id(url):
    m = re.search(r"(?:v=|youtu\.be/)([A-Za-z0-9_-]{11})", url)
    return m.group(1) if m else None

def fetch_youtube(url):
    vid = extract_youtube_id(url)
    if not vid: raise ValueError("유튜브 URL 인식 실패")
    title = f"YouTube: {vid}"
    try:
        r = requests.get(f"https://www.youtube.com/watch?v={vid}",
                         headers={"User-Agent":"Mozilla/5.0"}, timeout=10)
        m = re.search(r'"title":"([^"]+)"', r.text)
        if m: title = m.group(1)
    except: pass
    text = ""
    if YOUTUBE_OK:
        try:
            tlist = YouTubeTranscriptApi.list_transcripts(vid)
            try: t = tlist.find_transcript(["ko"])
            except: t = tlist.find_generated_transcript(["ko","en"])
            text = " ".join(s["text"] for s in t.fetch())
        except: text = f"[자막 추출 실패 - URL: {url}]"
    return title, text

def fetch_article(url):
    r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
    r.raise_for_status(); r.encoding = r.apparent_encoding
    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script","style","nav","footer","header","aside"]): tag.decompose()
    title = soup.title.string.strip() if soup.title else url
    text = "\n".join(l for l in soup.get_text(separator="\n",strip=True).splitlines() if l.strip())
    return title, text

def detect_and_fetch(url):
    if "youtube.com" in url or "youtu.be" in url: return fetch_youtube(url)
    return fetch_article(url)

def search_web(keyword, max_results=6):
    try:
        with DDGS() as d:
            return [{"title":r.get("title",""),"url":r.get("href",""),"snippet":r.get("body","")}
                    for r in d.text(keyword, max_results=max_results)]
    except: return []

def sources_path(policy): return POLICIES_DIR / policy / "sources.json"

def load_saved_sources(policy):
    p = sources_path(policy)
    if p.exists():
        try: return json.loads(p.read_text(encoding="utf-8"))
        except: pass
    return []

def save_sources(policy, sources):
    sources_path(policy).write_text(json.dumps(sources, ensure_ascii=False, indent=2), encoding="utf-8")

@st.cache_resource(show_spinner="PDF 읽는 중...")
def load_pdfs(policy):
    docs = {}
    for f in sorted((POLICIES_DIR/policy).glob("*.pdf")) + sorted((POLICIES_DIR/policy).glob("*.PDF")):
        try:
            parts = []
            with pdfplumber.open(f) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: parts.append(t)
            if parts: docs[f.name] = "\n".join(parts)
        except: pass
    return docs

def get_chunks(query, src_dict, max_chars=60000):
    kws = [w for w in query.lower().split() if len(w)>1]
    scored = sorted([(sum(tx.lower().count(k) for k in kws), nm, tx) for nm,tx in src_dict.items()], reverse=True)
    parts, total = [], 0
    for _, nm, tx in scored:
        chunk = f"[출처: {nm}]\n{tx[:8000]}\n"
        if total+len(chunk) > max_chars: break
        parts.append(chunk); total += len(chunk)
    return "\n---\n".join(parts)

def ask(query, context, history):
    sys = f"""당신은 정책 전문 리서처입니다. 아래 소스만 참고해 한국어로 답하세요.
출처 인용 필수. 없는 내용은 "확인되지 않습니다" 답변.
비교 질문은 표 사용.\n\n=== 소스 ===\n{context}"""
    hist = [{"role":"user" if m["role"]=="user" else "model","parts":[m["content"]]} for m in history[:-1]]
    return model.start_chat(history=hist).send_message(f"{sys}\n\n질문: {query}").text

def get_policies():
    if not POLICIES_DIR.exists(): return []
    return sorted([d.name for d in POLICIES_DIR.iterdir() if d.is_dir()])

import datetime
def now_str(): return datetime.datetime.now().strftime("%I:%M %p")


# ══════════════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════════════
policies = get_policies()
if not policies: st.warning("정책 폴더가 없습니다."); st.stop()

with st.sidebar:
    st.markdown("## 🏛️ 정책 공론장")
    st.caption("전남광주 통합특별시 시민주권")
    st.divider()

    selected_policy = st.radio("📂 정책 선택", policies,
                                format_func=lambda x: x.replace("-"," "))
    st.divider()

    # 관리자 로그인
    if "is_admin" not in st.session_state:
        st.session_state.is_admin = False

    if not st.session_state.is_admin:
        with st.expander("🔐 관리자 로그인"):
            pw = st.text_input("비밀번호", type="password", key="pw_input")
            if st.button("로그인", use_container_width=True):
                if pw == ADMIN_PASSWORD:
                    st.session_state.is_admin = True; st.rerun()
                else:
                    st.error("비밀번호가 틀렸습니다.")
    else:
        st.markdown('<div class="admin-badge">🔧 관리자 모드</div>', unsafe_allow_html=True)
        if st.button("로그아웃", use_container_width=True):
            st.session_state.is_admin = False; st.rerun()

# 정책 전환 시 초기화
if st.session_state.get("_cur") != selected_policy:
    st.session_state._cur        = selected_policy
    st.session_state.messages    = []
    st.session_state.web_sources = load_saved_sources(selected_policy)
    st.session_state.search_res  = []
    st.session_state.fc_idx      = 0; st.session_state.fc_show = False
    st.session_state.qz_idx      = 0; st.session_state.qz_score = 0
    st.session_state.qz_answered = False; st.session_state.qz_done = False
    st.session_state.report_text = ""

pdfs = load_pdfs(selected_policy)

# 체크 초기화
for f in pdfs:
    if f"ck_{f}" not in st.session_state: st.session_state[f"ck_{f}"] = True
for s in st.session_state.web_sources:
    if f"ck_{s['id']}" not in st.session_state: st.session_state[f"ck_{s['id']}"] = True


# ══════════════════════════════════════════════════════════════
# 메인
# ══════════════════════════════════════════════════════════════
st.markdown(f"## 🏛️ {selected_policy.replace('-',' ')}")

tab_chat, tab_studio = st.tabs(["💬 채팅", "🎓 스튜디오"])

# ──────────────────────────────────────────────────────────────
# 탭1: 채팅
# ──────────────────────────────────────────────────────────────
with tab_chat:
    src_col, chat_col = st.columns([1, 2.2], gap="large")

    # ── 소스 패널 ────────────────────────────────────────────
    with src_col:
        active_pdfs = [f for f in pdfs if st.session_state.get(f"ck_{f}", True)]
        active_web  = [s for s in st.session_state.web_sources if st.session_state.get(f"ck_{s['id']}", True)]
        total_active = len(active_pdfs) + len(active_web)

        st.markdown(f"### 📎 소스 ({total_active}개 선택)")

        c1, c2 = st.columns(2)
        if c1.button("전체 선택", use_container_width=True, key="sel"):
            for f in pdfs: st.session_state[f"ck_{f}"] = True
            for s in st.session_state.web_sources: st.session_state[f"ck_{s['id']}"] = True
            st.rerun()
        if c2.button("전체 해제", use_container_width=True, key="desel"):
            for f in pdfs: st.session_state[f"ck_{f}"] = False
            for s in st.session_state.web_sources: st.session_state[f"ck_{s['id']}"] = False
            st.rerun()

        st.markdown('<div class="src-header">📄 PDF 문서</div>', unsafe_allow_html=True)
        for fname in pdfs:
            st.checkbox(fname[:32]+("…" if len(fname)>32 else ""),
                        key=f"ck_{fname}", value=st.session_state.get(f"ck_{fname}", True))

        if st.session_state.web_sources:
            st.markdown('<div class="src-header">🌐 웹 & 유튜브</div>', unsafe_allow_html=True)
            for src in st.session_state.web_sources:
                icon = "▶️" if src.get("type")=="youtube" else "📰"
                cols = st.columns([5,1])
                cols[0].checkbox(f"{icon} {src['title'][:26]}{'…' if len(src['title'])>26 else ''}",
                                 key=f"ck_{src['id']}", value=st.session_state.get(f"ck_{src['id']}", True))
                # 관리자만 삭제 가능
                if st.session_state.is_admin:
                    if cols[1].button("✕", key=f"rm_{src['id']}"):
                        st.session_state.web_sources = [s for s in st.session_state.web_sources if s["id"]!=src["id"]]
                        save_sources(selected_policy, st.session_state.web_sources)
                        st.rerun()

        # ── 관리자 소스 추가 ────────────────────────────────
        if st.session_state.is_admin:
            st.divider()
            st.markdown("**🔧 소스 추가 (관리자)**")
            add_t1, add_t2 = st.tabs(["🔗 URL", "🔍 웹 리서치"])

            with add_t1:
                url_in = st.text_input("URL 입력", placeholder="https:// 또는 YouTube", key="url_in",
                                       label_visibility="collapsed")
                if st.button("추가", key="btn_url", use_container_width=True):
                    if url_in.startswith("http"):
                        with st.spinner("읽는 중..."):
                            try:
                                title, text = detect_and_fetch(url_in)
                                stype = "youtube" if ("youtube" in url_in or "youtu.be" in url_in) else "article"
                                ns = {"id":str(abs(hash(url_in)))[:10],"type":stype,
                                      "title":title,"url":url_in,"text":text[:20000]}
                                st.session_state.web_sources.append(ns)
                                st.session_state[f"ck_{ns['id']}"] = True
                                save_sources(selected_policy, st.session_state.web_sources)
                                st.success(f"추가: {title[:28]}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"실패: {e}")
                    else:
                        st.warning("URL을 입력하세요.")

            with add_t2:
                kw_in = st.text_input("검색 키워드", placeholder="예: 마을활동가 인정체계",
                                      key="kw_in", label_visibility="collapsed")
                if st.button("검색", key="btn_kw", use_container_width=True):
                    with st.spinner("검색 중..."):
                        st.session_state.search_res = search_web(kw_in)

                for i, r in enumerate(st.session_state.get("search_res", [])):
                    with st.container():
                        st.markdown(f"**{r['title'][:36]}**")
                        st.caption(f"{r['snippet'][:80]}…")
                        if st.button("＋ 소스 추가", key=f"add_{i}", use_container_width=True):
                            with st.spinner("추가 중..."):
                                try: title, text = fetch_article(r["url"])
                                except: title, text = r["title"], r["snippet"]
                                ns = {"id":str(abs(hash(r["url"])))[:10],"type":"article",
                                      "title":title,"url":r["url"],"text":text[:20000]}
                                st.session_state.web_sources.append(ns)
                                st.session_state[f"ck_{ns['id']}"] = True
                                save_sources(selected_policy, st.session_state.web_sources)
                                st.rerun()
                        st.divider()

    # ── 카카오톡 채팅 ────────────────────────────────────────
    with chat_col:
        if total_active == 0:
            st.warning("⚠️ 왼쪽에서 소스를 선택해주세요.")
        else:
            st.caption(f"📎 {total_active}개 소스 참고 중 · PDF {len(active_pdfs)} · 웹 {len(active_web)}")

        # 채팅 버블 렌더링
        bubbles_html = '<div class="chat-wrap" id="chat-wrap">'
        if not st.session_state.messages:
            bubbles_html += '''
            <div style="text-align:center;margin:auto;color:#fff;opacity:.7;font-size:.95rem;padding:40px 0">
            💬 아래에서 질문을 선택하거나 직접 입력해보세요
            </div>'''
        for msg in st.session_state.messages:
            t = msg.get("time","")
            if msg["role"] == "user":
                bubbles_html += f'''
                <div class="bubble-user">
                  <div class="time">{t}</div>
                  <div class="bubble">{msg["content"].replace(chr(10),"<br>")}</div>
                </div>'''
            else:
                content = msg["content"].replace(chr(10),"<br>")
                bubbles_html += f'''
                <div class="bubble-ai">
                  <div class="avatar">🏛️</div>
                  <div class="bubble-body">
                    <div class="sender">정책 리서처</div>
                    <div class="bubble">{content}</div>
                    <div class="time">{t}</div>
                  </div>
                </div>'''
        bubbles_html += '</div>'
        st.markdown(bubbles_html, unsafe_allow_html=True)

        # 예시 질문 (대화 없을 때)
        if not st.session_state.messages:
            st.markdown("**빠른 질문:**")
            examples = [
                "마을활동가 인정 방식을 지역별로 비교해주세요",
                "활동가 역량 기준은 어떻게 정의되나요?",
                "기회소득과 인정체계 연계 방안은?",
                "우수 사례와 정책 제언을 알려주세요",
            ]
            c1, c2 = st.columns(2)
            for i, q in enumerate(examples):
                if (c1 if i%2==0 else c2).button(q, key=f"ex{i}", use_container_width=True):
                    st.session_state.messages.append({"role":"user","content":q,"time":now_str()})
                    st.rerun()

        # 입력창
        if prompt := st.chat_input("메시지를 입력하세요..."):
            if total_active == 0:
                st.warning("소스를 먼저 선택해주세요.")
            else:
                st.session_state.messages.append({"role":"user","content":prompt,"time":now_str()})
                with st.spinner("답변 생성 중..."):
                    all_docs = {f: pdfs[f] for f in active_pdfs}
                    for s in active_web: all_docs[s["title"]] = s["text"]
                    ctx = get_chunks(prompt, all_docs)
                    ans = ask(prompt, ctx, st.session_state.messages)
                st.session_state.messages.append({"role":"assistant","content":ans,"time":now_str()})
                st.rerun()

        if st.session_state.messages:
            if st.button("🗑️ 대화 초기화", key="clr"):
                st.session_state.messages = []; st.rerun()


# ──────────────────────────────────────────────────────────────
# 탭2: 스튜디오
# ──────────────────────────────────────────────────────────────
with tab_studio:
    st.markdown("### 🎓 스튜디오")
    combined = "\n\n".join(f"[{fn}]\n{tx[:3000]}" for fn, tx in pdfs.items())

    @st.cache_data(show_spinner=False)
    def gen_summary(pol, c):
        t = call_gemini(f"""'{pol}' 문서 지역/기관별 핵심 요약. JSON만:
[{{"title":"지역(30자)","points":["핵심1","핵심2","핵심3"],"keyword":"키워드"}}]
최대 8개.\n{c[:40000]}""").strip()
        if "```" in t: t=t.split("```")[1]; t=t[4:] if t.startswith("json") else t
        try: return json.loads(t)
        except: return []

    @st.cache_data(show_spinner=False)
    def gen_info(pol, c):
        t = call_gemini(f"""정책 문서 인포그래픽 데이터. JSON만:
{{"metrics":[{{"label":"","value":"","unit":""}}],
"regions":[{{"name":"","approach":"","score":70}}],
"timeline":[{{"year":"","event":""}}],
"key_issues":[""]}}\n{c[:40000]}""").strip()
        if "```" in t: t=t.split("```")[1]; t=t[4:] if t.startswith("json") else t
        try: return json.loads(t)
        except: return {}

    @st.cache_data(show_spinner=False)
    def gen_mm(pol, c):
        t = call_gemini(f"""정책 마인드맵. Graphviz DOT만:
digraph {{graph[rankdir=LR] node[shape=rounded,style=filled,fontname="Helvetica"]
"중심" [fillcolor="#2E7D32",fontcolor=white]
"중심" -> "개념" ...}}\n{c[:30000]}""").strip()
        if "```" in t: t=t.split("```")[1]; t=t[3:] if t.startswith("dot") else t
        return t

    @st.cache_data(show_spinner=False)
    def gen_fc(pol, c):
        t = call_gemini(f"""'{pol}' 플래시카드 10개. JSON만:
[{{"question":"질문","answer":"답변(2~3문장)"}}]\n{c[:30000]}""").strip()
        if "```" in t: t=t.split("```")[1]; t=t[4:] if t.startswith("json") else t
        try: return json.loads(t)
        except: return []

    @st.cache_data(show_spinner=False)
    def gen_qz(pol, c):
        t = call_gemini(f"""'{pol}' 4지선다 8문제. JSON만:
[{{"question":"","options":["①","②","③","④"],"answer":0,"explanation":""}}]\n{c[:30000]}""").strip()
        if "```" in t: t=t.split("```")[1]; t=t[4:] if t.startswith("json") else t
        try: return json.loads(t)
        except: return []

    s1, s2, s3, s4, s5 = st.tabs(["📋 요약 카드","📊 인포그래픽","🗺️ 마인드맵","🃏 플래시카드","🧠 퀴즈 & 보고서"])

    with s1:
        if st.button("🔄 생성", key="gs"): gen_summary.clear(); st.rerun()
        with st.spinner("..."): cards = gen_summary(selected_policy, combined)
        if cards:
            cols = st.columns(2)
            for i,c in enumerate(cards):
                with cols[i%2]:
                    pts="".join(f"<li>{p}</li>" for p in c.get("points",[]))
                    kw=c.get("keyword","")
                    st.markdown(f"""<div style="background:#f8fdf8;border-left:4px solid #2E7D32;border-radius:8px;padding:14px 18px;margin-bottom:10px">
<b style="color:#1B5E20">{'🏷️ '+kw+' · ' if kw else ''}{c.get('title','')}</b>
<ul style="margin:8px 0 0;padding-left:18px;color:#333;font-size:.87rem;line-height:1.7">{pts}</ul></div>""",unsafe_allow_html=True)
        else: st.info("생성 버튼을 눌러주세요.")

    with s2:
        if st.button("🔄 생성", key="gi"): gen_info.clear(); st.rerun()
        with st.spinner("..."): info = gen_info(selected_policy, combined)
        if info:
            mets=info.get("metrics",[])
            if mets:
                st.markdown("#### 📌 주요 수치")
                mc=st.columns(min(len(mets),3))
                for i,m in enumerate(mets[:6]):
                    with mc[i%3]:
                        st.markdown(f'<div class="metric-box"><div class="num">{m.get("value","–")}<span style="font-size:.9rem;color:#555"> {m.get("unit","")}</span></div><div class="lbl">{m.get("label","")}</div></div>',unsafe_allow_html=True)
            regs=info.get("regions",[])
            if regs:
                st.markdown("#### 🗺️ 지역별 제도화 수준")
                for r in regs:
                    sc=int(r.get("score",50))
                    st.markdown(f"""<div style="background:#fff;border:1px solid #C8E6C9;border-radius:8px;padding:12px 16px;margin-bottom:6px">
<b style="color:#1B5E20">📍 {r.get('name','')}</b>
<div style="font-size:.82rem;color:#555;margin:4px 0 6px">{r.get('approach','')}</div>
<div style="background:#E8F5E9;border-radius:4px;height:10px"><div style="background:#2E7D32;border-radius:4px;height:10px;width:{sc}%"></div></div>
<div style="text-align:right;font-size:.75rem;color:#2E7D32;margin-top:2px">{sc}점</div></div>""",unsafe_allow_html=True)
            tl=info.get("timeline",[])
            if tl:
                st.markdown("#### 📅 주요 연혁")
                for item in tl:
                    a,b=st.columns([1,5]); a.markdown(f"**{item.get('year','')}**"); b.markdown(item.get("event",""))
            issues=info.get("key_issues",[])
            if issues:
                st.markdown("#### ⚡ 핵심 과제")
                ic=st.columns(2)
                for i,iss in enumerate(issues): ic[i%2].markdown(f"- {iss}")
        else: st.info("생성 버튼을 눌러주세요.")

    with s3:
        if st.button("🔄 생성", key="gm"): gen_mm.clear(); st.rerun()
        with st.spinner("..."): dot = gen_mm(selected_policy, combined)
        if dot:
            try: st.graphviz_chart(dot, use_container_width=True)
            except Exception as e: st.error(str(e)); st.code(dot)
        else: st.info("생성 버튼을 눌러주세요.")

    with s4:
        if st.button("🔄 생성", key="gf"): gen_fc.clear(); st.session_state.fc_idx=0; st.session_state.fc_show=False; st.rerun()
        with st.spinner("..."): fc = gen_fc(selected_policy, combined)
        if fc:
            n=len(fc); idx=st.session_state.fc_idx%n; card=fc[idx]
            st.markdown(f"**{idx+1} / {n}**")
            if not st.session_state.fc_show:
                st.markdown(f'<div class="flashcard">❓ {card["question"]}</div>',unsafe_allow_html=True)
                if st.button("답 보기 👁️",use_container_width=True): st.session_state.fc_show=True; st.rerun()
            else:
                st.markdown(f'<div class="flashcard">💡 {card["answer"]}</div>',unsafe_allow_html=True)
                c1,c2=st.columns(2)
                if c1.button("⬅️",use_container_width=True): st.session_state.fc_idx=(idx-1)%n; st.session_state.fc_show=False; st.rerun()
                if c2.button("➡️",use_container_width=True): st.session_state.fc_idx=(idx+1)%n; st.session_state.fc_show=False; st.rerun()
        else: st.info("생성 버튼을 눌러주세요.")

    with s5:
        ql, qr = st.columns(2)
        with ql:
            st.markdown("#### 🧠 퀴즈")
            if st.button("🔄 생성", key="gq"): gen_qz.clear(); st.session_state.qz_idx=0; st.session_state.qz_score=0; st.session_state.qz_answered=False; st.session_state.qz_done=False; st.rerun()
            with st.spinner("..."): qz = gen_qz(selected_policy, combined)
            if qz and not st.session_state.qz_done:
                qi=st.session_state.qz_idx
                if qi<len(qz):
                    q=qz[qi]
                    st.markdown(f"**{qi+1}/{len(qz)}** | 점수: {st.session_state.qz_score}")
                    st.markdown(f"**{q['question']}**")
                    for oi,opt in enumerate(q["options"]):
                        if st.button(opt,key=f"o{qi}{oi}",use_container_width=True,disabled=st.session_state.qz_answered):
                            st.session_state.qz_answered=True
                            if oi==q["answer"]: st.session_state.qz_score+=1; st.success("✅ 정답!")
                            else: st.error(f"❌ 정답: {q['options'][q['answer']]}")
                            st.info(f"📖 {q.get('explanation','')}")
                    if st.session_state.qz_answered:
                        if st.button("다음 ➡️",use_container_width=True,key=f"qn{qi}"):
                            st.session_state.qz_idx+=1; st.session_state.qz_answered=False
                            if st.session_state.qz_idx>=len(qz): st.session_state.qz_done=True
                            st.rerun()
                else: st.session_state.qz_done=True; st.rerun()
            elif st.session_state.qz_done and qz:
                pct=int(st.session_state.qz_score/len(qz)*100)
                st.markdown(f"## 🎉 {st.session_state.qz_score}/{len(qz)} ({pct}점)")
                if pct>=80: st.success("우수!")
                elif pct>=50: st.warning("복습이 필요합니다.")
                else: st.error("요약 카드부터 다시 시작하세요.")
                if st.button("다시 도전"): st.session_state.qz_idx=0; st.session_state.qz_score=0; st.session_state.qz_answered=False; st.session_state.qz_done=False; st.rerun()
            else: st.info("생성 버튼을 눌러주세요.")

        with qr:
            st.markdown("#### 📄 보고서 & PPT")
            if st.button("📝 보고서 생성", type="primary", use_container_width=True):
                with st.spinner("보고서 작성 중... (약 1분)"):
                    full="\n\n".join(f"[{fn}]\n{tx[:5000]}" for fn,tx in pdfs.items())
                    st.session_state.report_text = call_gemini(f"""아래 문서로 종합 보고서. 마크다운:
# {selected_policy} 종합 분석 보고서
## 1. 개요 ## 2. 지역별 현황 ## 3. 주요 쟁점 ## 4. 우수 사례 ## 5. 정책 제언
각 섹션 구체적으로.\n{full[:50000]}""")
            if st.session_state.report_text:
                st.markdown(st.session_state.report_text[:2000]+"…")
                st.download_button("📥 보고서 (.md)", data=st.session_state.report_text,
                    file_name=f"{selected_policy}_보고서.md", mime="text/markdown", use_container_width=True)
