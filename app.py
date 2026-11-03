import io, json, re, time
import pandas as pd
import requests
import streamlit as st
from rapidfuzz import fuzz
from google import genai
from google.genai import types

st.set_page_config(page_title="Survey Validation Studio", page_icon="◈", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1.4rem;max-width:1500px}
[data-testid="stSidebar"]{border-right:1px solid rgba(120,130,160,.2)}
[data-testid="stMetric"]{background:rgba(100,110,170,.08);padding:14px;border-radius:14px}
[data-testid="stFileUploader"]{border:1px dashed rgba(100,110,160,.5);border-radius:14px;padding:10px}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:600}
h1{letter-spacing:-.035em}
</style>""", unsafe_allow_html=True)

PROVIDERS=["Gemini","Groq","OpenRouter","Ollama"]
DEFAULT_MODELS={
"Gemini":["gemini-2.5-flash","gemini-2.5-pro"],
"Groq":["llama-3.3-70b-versatile","openai/gpt-oss-120b","llama-3.1-8b-instant"],
"OpenRouter":["openrouter/free"],
"Ollama":["qwen3:8b","llama3.1:8b","gemma3:4b"]}
BASE={"Groq":"https://api.groq.com/openai/v1","OpenRouter":"https://openrouter.ai/api/v1","Ollama":"http://localhost:11434/v1"}
BRAND_LABELS=["VALID_LOCAL","VALID_GLOBAL","INVALID","GARBAGE","NEEDS_REVIEW"]
TEXT_LABELS=["VALID","INVALID","GARBAGE","NON_INSIGHTFUL","NEEDS_REVIEW"]

def clean(x):
    if x is None: return ""
    try:
        if pd.isna(x): return ""
    except Exception: pass
    return str(x).strip()

def norm(x): return re.sub(r"\s+"," ",re.sub(r"[^\w\s]"," ",clean(x).lower())).strip()

def extract_json(s):
    s=(s or "").strip()
    s=re.sub(r"^```(?:json)?\s*|\s*```$","",s,flags=re.I)
    try:return json.loads(s)
    except Exception:
        a=s.find("{"); b=s.rfind("}")
        if a>=0 and b>a:return json.loads(s[a:b+1])
        raise ValueError("Model did not return valid JSON.")

def api_key(p): return st.session_state.get("key_"+p,"").strip()

def model_catalogue(p,key="",ollama_url="http://localhost:11434"):
    try:
        if p=="Gemini" and key:
            c=genai.Client(api_key=key)
            names=[m.name.replace("models/","") for m in c.models.list() if "gemini" in (m.name or "").lower()]
            return names or DEFAULT_MODELS[p]
        if p in ("Groq","OpenRouter") and key:
            r=requests.get(BASE[p]+"/models",headers={"Authorization":"Bearer "+key},timeout=15);r.raise_for_status()
            ids=[x["id"] for x in r.json().get("data",[]) if x.get("id")]
            if p=="OpenRouter":
                free=[x for x in ids if x.endswith(":free") or x=="openrouter/free"]
                return free or ids or DEFAULT_MODELS[p]
            return ids or DEFAULT_MODELS[p]
        if p=="Ollama":
            url=ollama_url.rstrip("/").replace("/v1","")+"/api/tags"
            r=requests.get(url,timeout=5);r.raise_for_status()
            ids=[x["name"] for x in r.json().get("models",[]) if x.get("name")]
            return ids or DEFAULT_MODELS[p]
    except Exception: pass
    return DEFAULT_MODELS[p]

def call_model(p,model,prompt):
    key=api_key(p)
    if p=="Gemini":
        if not key: raise RuntimeError("Gemini API key missing")
        c=genai.Client(api_key=key)
        r=c.models.generate_content(model=model,contents=prompt,config=types.GenerateContentConfig(response_mime_type="application/json",temperature=0.1))
        return extract_json(r.text or "")
    if p=="Ollama" and not requests.get("http://localhost:11434/api/tags",timeout=3).ok:
        raise RuntimeError("Ollama server is not running")
    if p!="Ollama" and not key: raise RuntimeError(p+" API key missing")
    headers={"Content-Type":"application/json"}
    if key:headers["Authorization"]="Bearer "+key
    if p=="OpenRouter":
        headers["HTTP-Referer"]="https://localhost";headers["X-Title"]="Survey Validation Studio"
    payload={"model":model,"messages":[{"role":"system","content":"Return valid JSON only."},{"role":"user","content":prompt}],"temperature":0.1}
    if p!="Ollama":payload["response_format"]={"type":"json_object"}
    url=BASE[p]+"/chat/completions"
    r=requests.post(url,headers=headers,json=payload,timeout=120)
    if r.status_code==400 and "response_format" in r.text:
        payload.pop("response_format",None);r=requests.post(url,headers=headers,json=payload,timeout=120)
    r.raise_for_status()
    return extract_json(r.json()["choices"][0]["message"]["content"])

def route(prompt,primary,model,fallbacks,fbmodels):
    errors=[]
    for p in [primary]+[x for x in fallbacks if x!=primary]:
        m=model if p==primary else fbmodels.get(p,DEFAULT_MODELS[p][0])
        try:return call_model(p,m,prompt),p,m
        except Exception as e:errors.append(p+": "+str(e)[:180])
    raise RuntimeError("All providers failed. "+" | ".join(errors))

def read_data(f):
    return pd.read_csv(f) if f.name.lower().endswith(".csv") else pd.read_excel(f)

def excel_file(sheets):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        for name,df in sheets.items():
            df.to_excel(w,index=False,sheet_name=name[:31] or "Report")
            ws=w.sheets[name[:31] or "Report"];ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width=min(46,max(12,max((len(str(c.value)) for c in col if c.value is not None),default=10)+2))
    return b.getvalue()

def provider_panel():
    st.sidebar.markdown("## ◈ PROVIDER MANAGER")
    st.sidebar.caption("Credentials are held in this browser session.")
    for p in PROVIDERS:
        if p=="Ollama":st.sidebar.text_input("Ollama local URL", "http://localhost:11434",key="ollama_url")
        else:st.sidebar.text_input(p+" API key",type="password",key="key_"+p)
    if st.sidebar.button("Test configured providers",use_container_width=True):
        results=[]
        for p in PROVIDERS:
            if p!="Ollama" and not api_key(p):results.append((p,"Not configured"));continue
            try:
                ms=model_catalogue(p,api_key(p),st.session_state.get("ollama_url","http://localhost:11434"))
                if p=="Ollama":
                    url=st.session_state.get("ollama_url","http://localhost:11434").rstrip("/")+"/api/tags"
                    rr=requests.get(url,timeout=5);rr.raise_for_status()
                    results.append((p,"Connected — "+str(len(rr.json().get("models",[])))+" local model(s)"))
                else:
                    call_model(p,ms[0],'Return JSON {"connection":"ok"}')
                    results.append((p,"Connected — "+ms[0]))
            except Exception as e:results.append((p,"Failed — "+str(e)[:150]))
        st.session_state["provider_results"]=results
    for p,s in st.session_state.get("provider_results",[]):st.sidebar.caption(p+": "+s)

def sidebar_model_config():
    p=st.sidebar.selectbox("Primary AI platform",PROVIDERS,key="primary_provider")
    url=st.session_state.get("ollama_url","http://localhost:11434")
    models=model_catalogue(p,api_key(p),url)
    model=st.sidebar.selectbox("Primary model",models,key="primary_model")
    batch=int(st.sidebar.number_input("Batch size",1,1000,50,step=10))
    usefb=st.sidebar.checkbox("Automatic provider fallback",True)
    fbs=st.sidebar.multiselect("Fallback order",[x for x in PROVIDERS if x!=p],default=[x for x in ["Groq","OpenRouter","Ollama"] if x!=p]) if usefb else []
    fbmodels={}
    for fp in fbs:
        ms=model_catalogue(fp,api_key(fp),url)
        fbmodels[fp]=st.sidebar.selectbox("Fallback model — "+fp,ms,key="fb_"+fp)
    return p,model,fbs,fbmodels,batch

def classify_rows(kind,question,df,col,batch,p,model,fbs,fbmodels,country_col=None,fixed_country=""):
    uniques=[];seen=set();keys=[]
    for _,r in df.iterrows():
        val=clean(r[col]);country=clean(r[country_col]) if country_col else fixed_country
        k=(norm(val),norm(country) if kind=="brand" else "")
        keys.append(k)
        if val and k not in seen:seen.add(k);uniques.append((val,country,k))
    cache={};progress=st.progress(0);msg=st.empty()
    for start in range(0,len(uniques),batch):
        chunk=uniques[start:start+batch]
        rows=[{"index":i+1,"response":v,"country":c} for i,(v,c,k) in enumerate(chunk)]
        if kind=="brand":
            labels=BRAND_LABELS
            instruction="Validate whether each answer is a real brand/company relevant to the question and market. Distinguish local/global brands, invalid answers, gibberish and ambiguous cases."
            schema='{"results":[{"index":1,"classification":"VALID_LOCAL","confidence":90,"normalized":"","reason":""}]}'
        else:
            labels=TEXT_LABELS
            instruction="Classify meaningful relevant answers as VALID including negative opinions; INVALID means understandable but irrelevant; GARBAGE is gibberish; NON_INSIGHTFUL is a non-answer; ambiguous is NEEDS_REVIEW."
            schema='{"results":[{"index":1,"classification":"VALID","confidence":90,"reason":""}]}'
        prompt=f"{instruction}\nQuestion: {question}\nAllowed labels: {labels}\nReturn one result for every index. Confidence 0-100. Input: {json.dumps(rows,ensure_ascii=False)}\nJSON schema: {schema}"
        try:
            data,usedp,usedm=route(prompt,p,model,fbs,fbmodels)
            got={int(x["index"]):x for x in data.get("results",[]) if str(x.get("index","")).isdigit()}
            for i,(v,c,k) in enumerate(chunk,1):
                x=got.get(i,{})
                label=str(x.get("classification","NEEDS_REVIEW")).upper()
                if label not in labels:label="NEEDS_REVIEW"
                conf=float(x.get("confidence",0) or 0)
                if conf<55:label="NEEDS_REVIEW"
                cache[k]={"Status":label,"Confidence":conf,"Reason":x.get("reason",""),"Normalized":x.get("normalized",""),"Provider":usedp,"Model":usedm}
        except Exception as e:
            for v,c,k in chunk:cache[k]={"Status":"NEEDS_REVIEW","Confidence":0,"Reason":str(e),"Normalized":"","Provider":"FAILED","Model":""}
        progress.progress(min(1,(start+len(chunk))/max(1,len(uniques))))
        msg.info(f"Processed {min(start+len(chunk),len(uniques)):,} / {len(uniques):,} unique answers")
    return pd.DataFrame([cache.get(k,{"Status":"NEEDS_REVIEW","Confidence":0,"Reason":"Blank or missing","Normalized":"","Provider":"","Model":""}) for k in keys],index=df.index)

def questionnaire_text(file):
    name=file.name.lower()
    if name.endswith(".docx"):
        from docx import Document
        doc=Document(file)
        parts=[p.text for p in doc.paragraphs]
        for t in doc.tables:parts.extend(" | ".join(c.text for c in r.cells) for r in t.rows)
        return "\n".join(parts)
    if name.endswith(".pdf"):
        import pypdf
        return "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(file).pages)
    return file.getvalue().decode("utf-8",errors="replace")

def contradiction_page(p,model,fbs,fbmodels,batch):
    st.title("🧩 Questionnaire Contradiction Checker")
    st.write("Upload a questionnaire and respondent dataset. AI proposes checks; you review and approve them before respondent-level execution.")
    c1,c2=st.columns(2)
    qf=c1.file_uploader("Questionnaire (DOCX / PDF / TXT)",["docx","pdf","txt","md"],key="contr_q")
    df_file=c2.file_uploader("Respondent dataset (Excel / CSV)",["xlsx","xls","csv"],key="contr_data")
    if not qf or not df_file:
        st.info("Upload both files to start.")
        return
    try:qtext=questionnaire_text(qf);data=read_data(df_file)
    except Exception as e:st.error("Unable to read files: "+str(e));return
    st.success(f"Loaded {len(data):,} rows and {len(data.columns)} columns.")
    idcol=st.selectbox("Select the unique respondent ID column (required each time)",list(data.columns),key="contr_id")
    if data[idcol].isna().any() or data[idcol].duplicated().any():
        st.error("The selected ID column has blanks or duplicates. Choose a genuinely unique, populated respondent ID column.")
        return
    if st.button("Analyze questionnaire and generate checks",type="primary",use_container_width=True):
        with st.spinner("Analyzing questionnaire..."):
            prompt=f"""Act as an experienced market-research questionnaire auditor. Identify logical contradictions and suspicious patterns across questions, considering skip logic, routing, multi-select, DK/Prefer not to say, and missing values. Do not call unusual-but-possible answers contradictions. Do not invent variable names. Return JSON only: {{"checks":[{{"check_id":"C001","check_name":"Short name","description":"Why suspicious","variables":["question variable names or labels"],"logic":"Precise plain-language condition","severity":"High|Medium|Low","type":"contradiction|suspicious_pattern"}}]}}. Questionnaire:\n{qtext[:40000]}"""
            try:
                result,up,um=route(prompt,p,model,fbs,fbmodels)
                st.session_state["contr_checks"]=result.get("checks",[])
                st.session_state["contr_dataset"]=data
                st.session_state["contr_idcol"]=idcol
                st.session_state["contr_questionnaire"]=qtext
                st.session_state["contr_provider"]=up+"/"+um
            except Exception as e:st.error("Questionnaire analysis failed: "+str(e))
    checks=st.session_state.get("contr_checks",[])
    if checks:
        st.success(f"{len(checks)} proposed checks generated using {st.session_state.get('contr_provider','AI')}. Review before execution.")
        st.warning("AI proposals are not automatically treated as truth. Confirm the rule and variable mapping.")
        approved=[]
        for i,ch in enumerate(checks):
            with st.expander(f"{ch.get('check_id','C'+str(i+1))} · {ch.get('check_name','Proposed check')}",expanded=i==0):
                ch["check_name"]=st.text_input("Check name",ch.get("check_name",""),key=f"cn{i}")
                ch["description"]=st.text_area("Description",ch.get("description",""),key=f"cd{i}")
                ch["logic"]=st.text_area("Check logic",ch.get("logic",""),key=f"cl{i}")
                ch["severity"]=st.selectbox("Severity",["High","Medium","Low"],index=["High","Medium","Low"].index(ch.get("severity","Medium")) if ch.get("severity") in ["High","Medium","Low"] else 1,key=f"cs{i}")
                defaults=[v for v in ch.get("variables",[]) if v in data.columns]
                ch["mapped_variables"]=st.multiselect("Map to exact dataset columns",list(data.columns),default=defaults,key=f"cm{i}")
                ch["approved"]=st.checkbox("Approve this check",value=ch.get("approved",True),key=f"ca{i}")
                approved.append(ch)
        st.session_state["contr_checks"]=approved
        run=[x for x in approved if x.get("approved") and x.get("mapped_variables")]
        st.caption(f"{len(run)} approved checks ready to run.")
        if st.button("Run approved checks",type="primary",disabled=not run,use_container_width=True):
            all_flags={};audit=[];prog=st.progress(0);msg=st.empty()
            for ci,ch in enumerate(run,1):
                cols=ch["mapped_variables"];flagged=[]
                for start in range(0,len(data),batch):
                    part=data.iloc[start:start+batch]
                    rows=[{"row_index":int(ix),"respondent_id":str(row[idcol]),"answers":{c:clean(row[c]) for c in cols}} for ix,row in part.iterrows()]
                    prompt=f"""Evaluate only the approved rule below. Flag clear matches; do not flag missing values unless explicitly required. Return JSON: {{"results":[{{"row_index":0,"flag":true,"reason":"short evidence-based reason","observed_answers":{{"column":"value"}}}}]}}. Approved check: {json.dumps(ch,ensure_ascii=False)}. Respondent rows: {json.dumps(rows,ensure_ascii=False,default=str)}"""
                    try:
                        out,_,_=route(prompt,p,model,fbs,fbmodels)
                        got={int(x["row_index"]):x for x in out.get("results",[]) if str(x.get("row_index","")).lstrip("-").isdigit()}
                        for ix,row in part.iterrows():
                            x=got.get(int(ix),{})
                            if x.get("flag") is True:
                                flagged.append({idcol:row[idcol],"Check ID":ch.get("check_id"),"Check Name":ch.get("check_name"),"Description":ch.get("description"),"Reason":x.get("reason",""),"Observed Answers":json.dumps(x.get("observed_answers",{}),ensure_ascii=False,default=str),**{c:row[c] for c in cols}})
                    except Exception as e:
                        audit.append({"Check ID":ch.get("check_id"),"Check":ch.get("check_name"),"Execution status":"ERROR: "+str(e)[:200],"Variables":", ".join(cols),"Logic":ch.get("logic")})
                    msg.info(f"Check {ci}/{len(run)} · {ch.get('check_name')} · {min(start+len(part),len(data)):,}/{len(data):,} respondents")
                all_flags[ch.get("check_id","Check")]=pd.DataFrame(flagged)
                audit.append({"Check ID":ch.get("check_id"),"Check":ch.get("check_name"),"Execution status":f"Completed — {len(flagged)} flagged","Variables":", ".join(cols),"Logic":ch.get("logic")})
                prog.progress(ci/len(run))
            st.session_state["contr_flags"]=all_flags;st.session_state["contr_audit"]=pd.DataFrame(audit);st.session_state["contr_idcol"]=idcol
    flags=st.session_state.get("contr_flags")
    if flags is not None:
        st.subheader("Flagged respondents")
        unique_ids=set()
        sheets={}
        for check in st.session_state.get("contr_checks",[]):
            key=check.get("check_id","")
            frame=flags.get(key,pd.DataFrame())
            if frame.empty:continue
            unique_ids.update(frame[st.session_state["contr_idcol"]].astype(str))
            sheets=(sheets | {str(check.get("check_name",key))[:31]:frame})
            with st.expander(f"{check.get('check_name',key)} · {len(frame)} flagged"):
                st.write(check.get("description",""));st.dataframe(frame,use_container_width=True)
        a,b=st.columns(2);a.metric("Checks run",len(flags));b.metric("Unique respondents flagged",len(unique_ids))
        if not sheets:st.info("No respondents were flagged by the approved checks.")
        sheets["Check Audit"]=st.session_state.get("contr_audit",pd.DataFrame())
        st.download_button("Download grouped contradiction Excel report",excel_file(sheets),"questionnaire_contradiction_report.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

def main():
    st.sidebar.markdown("# ◈ VALIDATION STUDIO")
    st.sidebar.caption("Survey data quality workspace")
    provider_panel()
    p,model,fbs,fbmodels,batch=sidebar_model_config()
    page=st.sidebar.radio("WORKSPACES",["Dashboard","Brand Validation","Text Response Validation","Contradiction Checker"])
    if page=="Dashboard":
        st.title("Survey Validation Studio")
        st.caption("One workspace for brand quality, open-end quality and questionnaire consistency.")
        a,b,c=st.columns(3);a.metric("AI providers","4","Gemini · Groq · OpenRouter · Ollama");b.metric("Validation modules","3");c.metric("Batch size",f"{batch:,}","Configurable")
        x,y,z=st.columns(3)
        x.markdown("### 🏷️ Brand Validation\nValidate brand lists by market, classify local/global brands and review confidence.")
        y.markdown("### 📝 Text Response Validation\nClassify open ends, detect suspicious patterns and near-duplicates.")
        z.markdown("### 🧩 Contradiction Checker\nAnalyze questionnaire, approve checks and flag respondent IDs.")
        st.info("Select a workspace from the left navigation. Configure provider keys above.")
    elif page=="Contradiction Checker":contradiction_page(p,model,fbs,fbmodels,batch)
    elif page=="Brand Validation":
        st.title("🏷️ Brand Validation")
        f=st.file_uploader("Upload brand dataset",["xlsx","xls","csv"],key="brand_upload")
        if f:
            df=read_data(f);st.success(f"{len(df):,} rows loaded")
            with st.form("brand_form"):
                q=st.text_area("Survey question");col=st.selectbox("Brand response column",df.columns)
                cc=st.selectbox("Country column",["Use one country"]+list(df.columns))
                country=st.text_input("Country for all respondents") if cc=="Use one country" else ""
                go=st.form_submit_button("Validate brands",type="primary")
            if go and q.strip():
                out=classify_rows("brand",q,df,col,batch,p,model,fbs,fbmodels,None if cc=="Use one country" else cc,country)
                res=pd.concat([df,out.add_prefix("Brand_")],axis=1);st.dataframe(res,use_container_width=True)
                st.download_button("Download Brand Validation Excel",excel_file({"All Results":res,"Needs Review":res[res["Brand_Status"]=="NEEDS_REVIEW"]}),"brand_validation.xlsx")
    elif page=="Text Response Validation":
        st.title("📝 Text Response Validation")
        f=st.file_uploader("Upload open-end dataset",["xlsx","xls","csv"],key="text_upload")
        if f:
            df=read_data(f);st.success(f"{len(df):,} rows loaded")
            with st.form("text_form"):
                q=st.text_area("Survey question");col=st.selectbox("Text response column",df.columns)
                threshold=st.slider("Near-duplicate threshold",70,100,90)
                go=st.form_submit_button("Validate text responses",type="primary")
            if go and q.strip():
                out=classify_rows("text",q,df,col,batch,p,model,fbs,fbmodels)
                res=pd.concat([df,out.add_prefix("Text_")],axis=1)
                res["Suspicious_Flags"]=res[col].apply(lambda v:"; ".join(["Very short"] if len(clean(v))<3 else (["Repeated characters"] if re.search(r"(.)\\1{5,}",clean(v)) else [])))
                res["Duplicate_Flag"]=False;res["Duplicate_Of_Row"]=""
                seen=[]
                for ix,v in res[col].items():
                    nv=norm(v)
                    if not nv:continue
                    for oldix,ov in seen:
                        sim=100 if nv==ov else fuzz.ratio(nv,ov)
                        if sim>=threshold:res.at[ix,"Duplicate_Flag"]=True;res.at[ix,"Duplicate_Of_Row"]=str(oldix+2);break
                    seen.append((ix,nv))
                st.dataframe(res,use_container_width=True)
                st.download_button("Download OE Validation Excel",excel_file({"All Results":res,"Flagged":res[(res["Duplicate_Flag"])|(res["Suspicious_Flags"]!="")]}),"oe_validation.xlsx")

if __name__=="__main__":main()
