# check_ms.py — RESP 원고 검사기(HIMJ, 2026-11-21). 저널 한도 > 플레이북(3선 표·네모·지문) > 문체·수치. sanitize_docx.py 뒤에 돌린다. FAIL 0이면 PASS.
import os, re, subprocess, zipfile
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
import content_resp as C
from build_ms import read_draft, TITLE, ORCID, KEYS, AI_STATEMENT, GITHUB, FUNDING, COI, CONTRIB, IRB, OUT, REPO, ACK, statements
R = []
_tf0 = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "review", "_trace_terms_himj.txt"); _terms0 = [l.strip() for l in open(_tf0, encoding="utf-8") if l.strip()] if os.path.exists(_tf0) else []
def ok(c, m): R.append(("ok  " if c else "FAIL") + " " + m)
W = dict(l.split("=") for l in open(os.path.join(OUT, "wordcount.txt")).read().split())
D = read_draft(); AB = D["Abstract"]
MSF, SUPF, STF, CLF, AUF = "Manuscript.docx", "Supplementary_Material.docx", "STROBE_Checklist.docx", "Cover_Letter.docx", "Title_Page.docx"
for f in (MSF, SUPF, STF, CLF, AUF, "Table1.docx", "Table2.docx", "Figures/Figure1.tif", "Figures/Figure2.tif", "Figures/Figure3.tif", "Figures/Figure4.tif"): ok(os.path.exists(os.path.join(OUT, f)), f"파일 {f}")
# 1. HIMJ 한도
from build_ms import ABS_HEADS, _sortkey
ok(int(W["body_words"]) <= 5000, f"본문 {W['body_words']} ≤ 5,000(초록·참고문헌 제외)")
aw = sum(len(p.split()) for p in AB); ok(aw <= 350, f"초록 {aw}단어 ≤ 350(라벨 포함)")
ok([p.split(":")[0] for p in AB] == ABS_HEADS, "구조화 초록 6항목(Background/Objective/Method/Results/Conclusion/Implications for HIM practice)")
ok(TITLE.count(":") == 0 and "229" not in TITLE and "standardi" not in TITLE.lower() and "2011" not in TITLE and not any(re.search("(?i)" + t, TITLE) for t in _terms0), f"제목 {len(TITLE.split())}단어, 쌍점·시리즈 틀·이전 제목 없음")
ok(int(W["refs"]) == len(KEYS), f"참고문헌 {len(KEYS)}")
ok(not re.search(r"\[@", " ".join(AB)), "초록 인용 없음")
KW = [k.strip() for k in C.KEYWORDS.split(";")]; ok(len(KW) >= 4, f"MeSH 키워드 {len(KW)}개(≥4)")
ok(len(C.RUNNING) <= 60, "짧은 제목")
ok(any(p.startswith("Implications for health information management practice") for p in D["Discussion"]), "고찰에 보건정보관리 시사점 문단")
d = Document(os.path.join(OUT, MSF)); _ps = [p.text.strip() for p in d.paragraphs]; _dx = "\n".join(_ps)
ok(_ps[0] == TITLE and _ps[1] == "Abstract", "본문 파일 = 제목 + 초록부터")
ok(not any(t in _dx for t in ("kimsd@", ORCID, "94 Wausan-ro")), "본문 파일에 교신 연락처 없음(제목면 분리)")
_ord = ["Abstract", "Introduction", "Methods", "Results", "Discussion", "Conclusion", "Acknowledgments", "Author contributions", "Ethical approval and informed consent", "Declaration of conflicting interests", "Funding", "Data availability", "References", "Figure legends"]; pos = [_ps.index(k) if k in _ps else -1 for k in _ord]
ok(all(p >= 0 for p in pos) and pos == sorted(pos), f"원고 순서 {pos}")
for h, _ in statements(): ok(h in _ps, f"선언 소제목 '{h}'")
ok(_dx.count("Table 1.") == 0 and "Table 2." not in _dx.split("\nReferences\n")[-1], "본문 파일에 표 없음(표는 별도 Word 파일)")
for n in (1, 2):
    tt = Document(os.path.join(OUT, f"Table{n}.docx")); ok(len(tt.tables) == 1 and tt.tables[0].rows[0].cells[0].text.startswith(f"Table {n}.") or any(p.text.startswith(f"Table {n}.") for p in tt.paragraphs) or tt.tables[0].cell(0,0).text.startswith(f"Table {n}"), f"Table{n}.docx 제목 포함")
a = Document(os.path.join(OUT, AUF)); _ax = "\n".join(p.text for p in a.paragraphs)
for t in (TITLE, "Seongdae Kim PhD", ORCID, "kimsd@hongik.ac.kr", "Phone:", "Research Article", "Word count:", "Short title:"): ok(t in _ax, f"제목면 '{t[:40]}'")
ok(AI_STATEMENT.split(";")[0][:60] in ACK and IRB in _dx and FUNDING in _dx and COI in _dx and CONTRIB in _dx, "AI·IRB·연구비·이해충돌·기여 정본 문장")
ok("study received no specific funding" in _dx.split("\nMethods\n")[1][:600], "Methods 첫머리 연구비 역할 진술(유지)")
z = zipfile.ZipFile(os.path.join(OUT, MSF)); ok("<w:drawing" not in z.read("word/document.xml").decode(), "원고에 그림 삽입 없음(별도 파일)")
for k in (1, 2, 3, 4):
    im = Image.open(os.path.join(OUT, "Figures", f"Figure{k}.tif")); dpi = round(im.info.get("dpi", (0, 0))[0]); ok(im.mode == "RGB" and dpi >= 300 and getattr(im, "n_frames", 1) == 1, f"Figure{k}.tif RGB·{dpi} dpi(≥300)·단일 층")
ok(not re.search(r"\b(standardis|analys(e|ed|ing)\b|modell|ageing|behaviour|colour|centre|artefact|hospitalis|penalis|programme)", _dx[:_dx.find("\nReferences\n")], re.I), "미국식 철자(참고문헌 원제 제외)")
# 2. 표·문서 위생
for nm, f in (("원고", MSF), ("보충", SUPF), ("점검표", STF), ("커버", CLF), ("제목면", AUF), ("표1", "Table1.docx"), ("표2", "Table2.docx")):
    dd = Document(os.path.join(OUT, f))
    for i, t in enumerate(dd.tables):
        b = t._tbl.tblPr.find(qn("w:tblBorders"))
        ok(b is not None and all(b.find(qn(f"w:{s}")).get(qn("w:val")) == "nil" for s in ("insideV", "insideH", "left", "right")), f"{nm} 표{i+1} 3선")
        bad = [r.text for j, row in enumerate(t.rows) if j > 0 for c in row.cells if len({id(x._tc) for x in row.cells}) == len(row.cells) for p in c.paragraphs for r in p.runs if r.bold and r.text.strip()]
        ok(not bad, f"{nm} 표{i+1} 데이터 칸 볼드 없음 {bad[:2]}")
        ok(not [c.text for r in t.rows for c in r.cells if re.fullmatch(r"[\d.()\- ]+\*+", c.text.strip())], f"{nm} 표{i+1} 유의성 별표 없음")
        ok("w:shd" not in t._tbl.xml or 'w:fill="auto"' in t._tbl.xml or not re.search(r'w:fill="(?!auto|FFFFFF)', t._tbl.xml), f"{nm} 표{i+1} 음영 없음")
    x = subprocess.run(["unzip", "-p", os.path.join(OUT, f), "word/*.xml"], capture_output=True, text=True).stdout
    ok(len(re.findall(r"<w:keepNext[ />]|<w:keepLines[ />]|<w:pageBreakBefore[ />]", x)) == 0, f"{nm} Word 검은 네모 0건")
    zz = zipfile.ZipFile(os.path.join(OUT, f)); core = zz.read("docProps/core.xml").decode(); app = zz.read("docProps/app.xml").decode()
    ok("2013-12-23" not in core and not any(n.startswith("customXml/") for n in zz.namelist()) and "<Words>0<" not in app, f"{nm} python-docx 지문 없음(sanitize 후)")
    ok("<dc:creator>Seongdae Kim</dc:creator>" in core, f"{nm} 작성자 메타")
# 3. 문체·인용·흔적
body = re.sub(r"\s+", " ", _dx[_dx.find("\nIntroduction\n"):_dx.find("\nAcknowledgments\n")])
ok(not re.search(r"\b(we|our|us|We|Our)\b", body), "1인칭 없음")
supt = "\n".join(p.text for p in Document(os.path.join(OUT, SUPF)).paragraphs); cov = subprocess.run(["pdftotext", os.path.join(OUT, CLF.replace(".docx", ".pdf")), "-"], capture_output=True, text=True).stdout
stt = "\n".join(c.text for t in Document(os.path.join(OUT, STF)).tables for r in t.rows for c in r.cells)
alltxt = _dx + supt + cov + _ax + stt
ok("—" not in alltxt and not re.search(r"[가-힣]", alltxt), "em-dash·한글 없음")
_tf = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "review", "_trace_terms_himj.txt")   # 흔적 검사 목록(저장소 밖)
_terms = [l.strip() for l in open(_tf, encoding="utf-8") if l.strip()] if os.path.exists(_tf) else []
ok(not _terms or not re.search("(?i)" + "|".join(_terms), alltxt), "이전 저널·저장소·제목·버전 흔적 없음")
ok(not re.search(r"\bsignifican", _dx + supt, re.I), "'significant' 없음")
ok("−" in body, "마이너스 U+2212")
ok(not re.search(r"\[(@|web:)", _dx) and "{{" not in _dx, "미해결 인용 없음")
_nsx = re.sub(r"\([^()]*?(19|20)\d\d[a-z]?\)", "", "\n".join(p.text for p in d.paragraphs))   # 저자-연도 인용 제외
disc = _nsx[_nsx.find("\nDiscussion\n"):_nsx.find("\nAcknowledgments\n")]
_dn = re.findall(r".{15}\d.{8}", disc)[:3]; ok(not _dn, f"고찰·결론에 숫자 없음 {_dn}")
ok(not any(":" in x[0] for x in C.METHODS + C.RESULTS), "소제목 쌍점 없음")
refs = _dx[_dx.find("\nReferences\n"):_dx.find("\nFigure legends\n")]
ok(not re.search(r"github|zenodo", refs, re.I), "참고문헌에 GitHub·Zenodo 없음")
rl = [l for l in refs.split("\n") if re.match(r"^\S.*\((19|20)\d\d[a-z]?\) ", l)]; ok(len(rl) == len(KEYS), f"참고문헌 줄 {len(rl)}")
ok(all(("DOI:" in l or "PMID:" in l or "http" in l) for l in rl), "참고문헌마다 DOI·PMID·URL")
ok(not any(re.match(r"^\d+\. ", l) for l in refs.split("\n")), "번호 없는 Sage Harvard 목록")
ok(not any(re.search(r"^([^,(]+, ){3}[^,(]+, et al", l) for l in rl), "저자 3명 초과 시 3명 + et al.")
_cit = re.findall(r"\(([A-Z][^()]*?(?:19|20)\d\d[a-z]?)\)", body); ok(len(_cit) >= 20, f"본문 저자-연도 인용 {len(_cit)}건")
ok(not re.search(r"\[\d+(-\d+)?\]", body), "숫자 인용 없음")
ok(GITHUB in _dx and not any(re.search("(?i)" + t, REPO) for t in _terms0), "단일 익명: Data availability에 새 저장소, 이전 저장소 이름 아님")
ok("supplementary" in body, "보충 인용 'supplementary …'")
for k, (kind, n) in C.SUP.items(): ok(f"supplementary {kind} {n}" in body + supt, f"보충 {kind} {n} 본문 인용·존재")
ok("supplementary appendix A" in body, "보충 appendix A 본문 인용")
ok("zenodo.PENDING" not in _dx, "Zenodo DOI 채움")
open(os.path.join(OUT, "check_ms.txt"), "w").write("\n".join(R)); nf = sum(r.startswith("FAIL") for r in R)
print("\n".join(r for r in R if r.startswith("FAIL")) or "", f"\n{len(R)} checks, FAIL {nf}", "PASS" if nf == 0 else "")
