import pdfplumber, re, json, sys
pdf=pdfplumber.open(sys.argv[1])
out=[]
for pg in pdf.pages:
    ws=pg.extract_words()
    bad_tops={round(w['top']) for w in ws if w['text'] in ('TOTALES:','ATHLETES','TOTALES')}
    ws=[w for w in ws if round(w['top']) not in bad_tops and not w['text'].startswith('https://') and w['top']>(185 if pg.page_number==1 else 28) and w['top']<pg.height-25]
    codes=[w for w in ws if re.fullmatch(r'\d{5}',w['text']) and 270<w['x0']<290]
    codes.sort(key=lambda w:w['top'])
    for k,c in enumerate(codes):
        y=c['top']
        lo=(codes[k-1]['top']+y)/2 if k else y-25
        hi=(codes[k+1]['top']+y)/2 if k+1<len(codes) else y+25
        band=[w for w in ws if lo<=w['top']<hi]
        line=lambda xs: " ".join(w['text'] for w in sorted(xs,key=lambda w:(round(w['top']/3),w['x0'])))
        name=line([w for w in band if 150<=w['x0']<265])
        school=line([w for w in band if w['x0']<150 and w['text'] not in ('ARGENTINA','(ARG)','(ARG),')])
        cat=line([w for w in band if w['x0']>=290 and abs(w['top']-y)<6])
        out.append({"cod":c['text'],"name":name,"school":school.rstrip(','),"cat":cat})
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False,indent=0)
print(len(out))
