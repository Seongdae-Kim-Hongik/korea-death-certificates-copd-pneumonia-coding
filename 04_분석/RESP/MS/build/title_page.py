# title_page.py — RESP 제목면(HIMJ 2026-11-21, 심사자에게 보내지 않는 별도 파일): 제목·저자(학위)·소속·ORCID·교신 연락처·단어/표/그림 수·감사·선언.
import os, subprocess
import content_resp as C
from build_ms import TX, TITLE, ORCID, AFFIL, ADDRESS, ACK, CONTRIB, statements, OUT
W = dict(l.split("=") for l in open(os.path.join(OUT, "wordcount.txt")).read().split())
J = TX(cite="sup")
J.para(f"<b>{TITLE}</b>")
J.para(f"<b>Short title:</b> {C.RUNNING}")
J.para("Seongdae Kim PhD", before=12); J.para(AFFIL)
J.para(f"ORCID: https://orcid.org/{ORCID}")
J.para("<b>Correspondence:</b>", before=12); J.para(f"Seongdae Kim PhD, {ADDRESS}. Phone: +82-2-320-1114. Email: kimsd@hongik.ac.kr")
J.para(f"<b>Article type:</b> Research Article", before=12)
J.para(f"<b>Word count:</b> {int(W['body_words']):,} (Introduction to Conclusion, excluding abstract and references); abstract {W['abstract_words']} words; {W['tables']} tables; {W['figures']} figures; {W['refs']} references")
J.para("<b>Acknowledgments</b>", before=12); J.para(ACK)
J.para("<b>Author contributions</b>", before=6); J.para(CONTRIB)
J.para("<b>Declarations</b>", before=6)
for h, t in statements(): J.para(f"<i>{h}.</i> {t}")
p = os.path.join(OUT, "Title_Page.docx"); J.doc_title = "Title page"; J.save(p)
subprocess.run(["/opt/homebrew/bin/soffice", "--headless", "-env:UserInstallation=file:///tmp/lo_resp_tp_himj", "--convert-to", "pdf", "--outdir", OUT, p], check=True, capture_output=True)
print("title page pages", subprocess.run(["pdfinfo", p.replace(".docx", ".pdf")], capture_output=True, text=True).stdout.split("Pages:")[1].split()[0])
