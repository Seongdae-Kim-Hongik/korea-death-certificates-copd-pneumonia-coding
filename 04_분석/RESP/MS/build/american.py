# american.py — 미국식 철자 변환(Sage 미국식 표기): 빌더가 원고·표·범례 문자열을 출력 직전에 변환한다(참고문헌 서지는 원문 유지).
import re
STEMS = ["standardis", "harmonis", "summaris", "characteris", "categoris", "recognis", "emphasis", "minimis", "maximis", "prioritis", "utilis", "organis",
         "randomis", "generalis", "normalis", "visualis", "hospitalis", "urbanis", "stabilis", "specialis", "realis", "finalis", "authoris", "localis", "conceptualis",
         "operationalis", "penalis", "parameteris", "criticis", "mobilis", "optimis", "hypothesis", "residualis", "restandardis", "anonymis", "contextualis", "dichotomis", "linearis", "regularis"]
WORDS = {"analyse": "analyze", "analysed": "analyzed", "analysing": "analyzing", "modelled": "modeled", "modelling": "modeling", "modeller": "modeler",
         "centre": "center", "centres": "centers", "centred": "centered", "neighbour": "neighbor", "neighbours": "neighbors", "neighbourhood": "neighborhood", "neighbourhoods": "neighborhoods", "neighbouring": "neighboring",
         "behaviour": "behavior", "behaviours": "behaviors", "behavioural": "behavioral", "ageing": "aging", "colour": "color", "colours": "colors", "favour": "favor", "favours": "favors", "favoured": "favored", "favourable": "favorable",
         "labelled": "labeled", "labelling": "labeling", "travelled": "traveled", "grey": "gray", "fibre": "fiber", "programme": "program", "programmes": "programs",
         "haemoglobin": "hemoglobin", "anaemia": "anemia", "paediatric": "pediatric", "oedema": "edema", "equalled": "equaled", "acknowledgement": "acknowledgment", "acknowledgements": "acknowledgments",
         "judgement": "judgment", "artefact": "artifact", "artefacts": "artifacts", "artefactual": "artifactual", "tonne": "ton", "licence": "license", "defence": "defense", "practise": "practice", "towards": "toward", "whilst": "while", "amongst": "among"}
def amer(t):
    if not t: return t
    def stem(m):
        w = m.group(0); low = w.lower()
        for s in STEMS:
            if low.startswith(s):
                r = s[:-1] + "z" + low[len(s):]
                return r[0].upper() + r[1:] if w[0].isupper() else r
        return w
    t = re.sub(r"\b[A-Za-z]*is(?:e|ed|es|ing|ation|ations|er|ers)\b", stem, t)
    def word(m):
        w = m.group(0); r = WORDS.get(w.lower())
        if r is None: return w
        return r[0].upper() + r[1:] if w[0].isupper() else r
    return re.sub(r"\b[A-Za-z]+\b", word, t)
if __name__ == "__main__":
    print(amer("Rates were standardised, analysed and modelled; otherwise rises; hypothesised, penalised, residualised; analyses; artefactual; ageing; emphasis; Standardisation; towards."))
