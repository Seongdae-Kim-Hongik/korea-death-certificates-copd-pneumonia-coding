#!/bin/bash
# run_all.sh — RESP 전체 빌드(HIMJ, 2026-11-21). 그림은 figs_resp.py, 참고문헌 캐시 refs_resp.json 그대로(네트워크 없음).
# Zenodo DOI는 <논문폴더>/06_GitHub_Zenodo/zenodo_doi_himj.txt 한 줄을 읽는다(없으면 PENDING).
set -e; cd "$(dirname "$0")"; O=../out_himj; mkdir -p $O/Figures
python3 -W ignore figs_resp.py
python3 -W ignore build_ms.py
python3 -W ignore esm_resp.py
python3 -W ignore title_page.py
python3 pb_to_br.py $O/Manuscript.docx $O/Supplementary_Material.docx $O/Title_Page.docx $O/Table1.docx $O/Table2.docx
python3 sanitize_docx.py --no-anon $O/Manuscript.docx $O/Supplementary_Material.docx $O/Title_Page.docx $O/Table1.docx $O/Table2.docx
for f in Supplementary_Material Title_Page Table1 Table2; do /opt/homebrew/bin/soffice --headless -env:UserInstallation=file:///tmp/lo_resp_fin_himj --convert-to pdf --outdir $O $O/$f.docx >/dev/null 2>&1; done
./word_pdf.sh $O/Manuscript.docx $O/Manuscript.pdf
python3 -W ignore strobe_pages.py $O/Manuscript.pdf
python3 pb_to_br.py $O/STROBE_Checklist.docx; python3 sanitize_docx.py --no-anon $O/STROBE_Checklist.docx
/opt/homebrew/bin/soffice --headless -env:UserInstallation=file:///tmp/lo_resp_fin_himj --convert-to pdf --outdir $O $O/STROBE_Checklist.docx >/dev/null 2>&1
python3 -W ignore check_ms.py
