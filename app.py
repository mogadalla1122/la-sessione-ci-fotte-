# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
import random
from docx import Document
from openpyxl import load_workbook

st.set_page_config(page_title="Quiz da file", layout="centered")
st.title("Quiz dalle tue domande (CSV / Excel / DOCX)")
st.markdown("Carica un file con domande e opzioni. L'app prova a individuare la risposta corretta automaticamente (colonna 'correct', asterisco '*' o testo in grassetto per .docx / Excel).")

def parse_csv_bytes(b):
    return pd.read_csv(BytesIO(b))

def parse_excel_bytes(b):
    return pd.read_excel(BytesIO(b), sheet_name=0)

def build_items_from_df_with_correct_col(df, correct_col):
    col_map = {c.lower(): c for c in df.columns}
    qcol = None
    for name in ("question", "domanda", "q"):
        if name in col_map:
            qcol = col_map[name]; break
    if qcol is None:
        qcol = df.columns[0]
    option_cols = [c for c in df.columns if c not in (qcol, correct_col)]
    items = []
    for _, row in df.iterrows():
        q = "" if pd.isna(row[qcol]) else str(row[qcol])
        opts = [("" if pd.isna(row[c]) else str(row[c])) for c in option_cols]
        corr = row[correct_col]
        corr_idx = None
        if pd.notna(corr):
            cs = str(corr).strip()
            if len(cs)==1 and cs.isalpha():
                idx = ord(cs.lower())-ord('a')
                if 0 <= idx < len(opts):
                    corr_idx = idx
            else:
                for i,o in enumerate(opts):
                    if str(o).strip() == cs:
                        corr_idx = i
                        break
        items.append({"question": q, "options": opts, "correct_index": corr_idx})
    return items

def detect_from_df(df):
    df_cols = [c.lower() for c in df.columns]
    col_map = {c.lower(): c for c in df.columns}
    for name in ("correct", "correct_answer", "answer", "risposta"):
        if name in col_map:
            return build_items_from_df_with_correct_col(df, col_map[name])
    possible_option_cols = [c for c in df.columns if c.lower().startswith("option") or c.lower().startswith("opt")]
    if len(possible_option_cols) < 2:
        qcol = None
        for name in ("question", "domanda", "q"):
            if name in col_map:
                qcol = col_map[name]; break
        if qcol is None:
            qcol = df.columns[0]
        option_cols = [c for c in df.columns if c != qcol]
    else:
        option_cols = possible_option_cols
        qcol = None
        for name in ("question", "domanda", "q"):
            if name in col_map:
                qcol = col_map[name]; break
        if qcol is None:
            qcol = df.columns[0]
    items = []
    for _, row in df.iterrows():
        q = str(row[qcol]) if pd.notna(row[qcol]) else ""
        opts = []
        correct_index = None
        for i, c in enumerate(option_cols):
            val = "" if pd.isna(row[c]) else str(row[c])
            if val.startswith("*"):
                clean = val.lstrip("*").strip()
                opts.append(clean)
                correct_index = i
            else:
                opts.append(val)
        items.append({"question": q, "options": opts, "correct_index": correct_index})
    return items

def parse_docx_bytes(b):
    doc = Document(BytesIO(b))
    paras = [p for p in doc.paragraphs if p.text.strip()!=""]
    items = []
    i = 0
    while i < len(paras):
        question_para = paras[i]
        question = question_para.text.strip()
        opts = []
        corr = None
        j = i+1
        while j < len(paras) and len(opts) < 10:
            p = paras[j]
            text = p.text.strip()
            if not text:
                break
            is_bold = any((run.bold is True) for run in p.runs)
            opts.append(text)
            if is_bold:
                corr = len(opts)-1
            j += 1
            if j < len(paras) and paras[j].text.strip().endswith("?"):
                break
        items.append({"question": question, "options": opts, "correct_index": corr})
        i = j
    return items

def try_parse_file(uploaded):
    name = uploaded.name.lower()
    data = uploaded.read()
    try:
        if name.endswith(".csv") or name.endswith(".txt"):
            df = parse_csv_bytes(data)
            items = detect_from_df(df)
            return items, "csv"
        if name.endswith(".xls") or name.endswith(".xlsx"):
            df = parse_excel_bytes(data)
            items = detect_from_df(df)
            try:
                wb = load_workbook(filename=BytesIO(data), read_only=False, data_only=True)
                ws = wb[wb.sheetnames[0]]
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in next(ws.iter_rows(min_row=1, max_row=1))]
                header_map = {h.lower(): idx for idx,h in enumerate(headers)}
                qcol_idx = header_map.get("question", 0)
                corr_colname = None
                for namec in ("correct","correct_answer","answer"):
                    if namec in header_map:
                        corr_colname = namec; break
                option_cols_idx = [i for i,h in enumerate(headers) if i!=qcol_idx and (corr_colname is None or h.lower()!=corr_colname)]
                items_bold = []
                for r_idx,row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
                    q = row[qcol_idx].value if qcol_idx < len(row) else ""
                    opts = []
                    corr = None
                    for j, cidx in enumerate(option_cols_idx):
                        if cidx < len(row):
                            cell = row[cidx]
                            val = "" if cell.value is None else str(cell.value)
                            opts.append(val)
                            if cell.font and cell.font.bold:
                                corr = j
                    items_bold.append({"question": q, "options": opts, "correct_index": corr})
                if any(it['correct_index'] is not None for it in items_bold):
                    return items_bold, "excel"
            except Exception:
                pass
            return items, "excel"
        if name.endswith(".docx"):
            items = parse_docx_bytes(data)
            return items, "docx"
    except Exception as e:
        st.error(f"Errore durante il parsing: {e}")
    return [], "unknown"

uploaded = st.file_uploader("Carica il file (CSV, Excel .xls/.xlsx, o DOCX)", type=["csv","xls","xlsx","docx"])
if uploaded is not None:
    with st.spinner("Parsing file..."):
        items, fmt = try_parse_file(uploaded)
    if not items:
        st.warning("Nessuna domanda trovata automaticamente. Se vuoi, carica un file di esempio o dimmi come sono evidenziate le risposte.")
    else:
        st.success(f"Trovate {len(items)} domande (formato stimato: {fmt}).")
        for i,it in enumerate(items[:5]):
            st.markdown(f"**{i+1}. {it['question']}**")
            for idx,opt in enumerate(it['options']):
                marker = "✅" if it['correct_index'] == idx else ""
                st.write(f"- {chr(ord('A')+idx)}) {opt} {marker}")
        shuffle = st.checkbox("Mescola le domande", value=True)
        max_q = st.number_input("Numero massimo di domande (0 = tutte)", min_value=0, max_value=len(items), value=0, step=1)
        if st.button("Inizia il quiz"):
            if shuffle:
                random.shuffle(items)
            if max_q and max_q>0:
                items = items[:max_q]
            st.session_state['quiz_items'] = items
            st.session_state['q_index'] = 0
            st.session_state['answers'] = [None]*len(items)
            st.experimental_rerun()

if 'quiz_items' in st.session_state:
    items = st.session_state['quiz_items']
    idx = st.session_state['q_index']
    if idx < len(items):
        it = items[idx]
        st.markdown(f"### Domanda {idx+1} / {len(items)}")
        st.write(it['question'])
        cols = st.columns([1,3])
        with cols[0]:
            chosen = st.radio("Scegli la risposta", [chr(ord('A')+i) for i in range(len(it['options']))], key=f"choice_{idx}")
        with cols[1]:
            for i,opt in enumerate(it['options']):
                tag = " (rilevata corretta)" if it['correct_index'] == i else ""
                st.write(f"{chr(ord('A')+i)}) {opt}{tag}")
            if st.button("Segna scelta come corretta", key=f"setcorr_{idx}"):
                it['correct_index'] = ord(chosen.lower())-97
                st.session_state['quiz_items'][idx] = it
                st.success("Impostata come corretta.")
        if st.button("Prossima"):
            st.session_state['answers'][idx] = ord(chosen.lower())-97
            st.session_state['q_index'] += 1
            st.experimental_rerun()
    else:
        answers = st.session_state['answers']
        score = 0
        details = []
        for i,it in enumerate(items):
            corr = it.get('correct_index')
            usr = answers[i]
            correct_letter = chr(ord('A')+corr) if corr is not None else "?"
            usr_letter = chr(ord('A')+usr) if usr is not None else "—"
            ok = (corr is not None and usr == corr)
            if ok: score += 1
            details.append({"question": it['question'], "your": usr_letter, "correct": correct_letter, "ok": ok, "options": it['options']})
        st.markdown("## Risultato finale")
        st.write(f"Punteggio: **{score} / {len(items)}** — {score/len(items)*100:.1f}%")
        st.markdown("### Dettaglio domande sbagliate")
        for i,d in enumerate(details):
            if not d['ok']:
                st.markdown(f"**{i+1}. {d['question']}**")
                st.write(f"Risposta tua: {d['your']} — Corretta: {d['correct']}")
                for j,opt in enumerate(d['options']):
                    mark = "✅" if chr(ord('A')+j) == d['correct'] else ""
                    st.write(f"- {chr(ord('A')+j)}) {opt} {mark}")
        if st.button("Esporta risultati CSV"):
            out = []
            for i,d in enumerate(details):
                out.append({"question": d['question'], "your": d['your'], "correct": d['correct'], "ok": d['ok']})
            df_out = pd.DataFrame(out)
            st.download_button("Scarica CSV", df_out.to_csv(index=False).encode('utf-8'), "risultati_quiz.csv", "text/csv")
        if st.button("Ricomincia"):
            for k in ['quiz_items','q_index','answers']:
                if k in st.session_state: del st.session_state[k]
            st.experimental_rerun()
