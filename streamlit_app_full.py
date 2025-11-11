
import streamlit as st
import pandas as pd
import io, json
from datetime import datetime

st.set_page_config(page_title="Paket-Konfigurator - Institut für Holztechnologie", layout="wide")

def load_excel(path_or_file):
    try:
        if hasattr(path_or_file, "read"):
            df = pd.read_excel(path_or_file, sheet_name=0)
        else:
            df = pd.read_excel(path_or_file, sheet_name=0)
        return df
    except Exception as e:
        st.error(f"Fehler beim Laden der Excel-Datei: {e}")
        return pd.DataFrame()

# Sidebar / Left column: Data source + controls
left, right = st.columns([3,7])

with left:
    st.image("https://placehold.co/120x60?text=Logo", width=120)
    st.header("Datenquelle")
    uploaded = st.file_uploader("Miete.xlsx hochladen (ersetzt lokale Datei)", type=["xlsx","xls"])
    use_uploaded = False
    if uploaded is not None:
        df = load_excel(uploaded)
        use_uploaded = True
        st.success("Datei geladen (Upload).")
    else:
        try:
            df = load_excel("Miete.xlsx")
            if df.empty:
                st.warning("Keine lokale Miete.xlsx gefunden. Bitte Datei hochladen.")
        except FileNotFoundError:
            df = pd.DataFrame()
            st.warning("Keine lokale Miete.xlsx gefunden. Bitte Datei hochladen.")
    st.markdown("---")
    st.subheader("Suche / Filter")
    search = st.text_input("Suchbegriff (Name)", value="")
    st.markdown("**Artikel-Liste**")
    if not df.empty:
        name_col = df.columns[0]
        price_col = df.columns[-1]
        visible = df[df[name_col].str.contains(search, case=False, na=False)] if search else df
        options = visible.apply(lambda r: f\"{r[name_col]} — {r[price_col]}\", axis=1).tolist()
        mapping = {opt: visible.iloc[i][name_col] for i,opt in enumerate(options)}
        st.text(f\"{len(options)} Artikel gefunden\")
        select_all = st.button(\"Alle auswählen\")
        clear_all = st.button(\"Auswahl löschen\")
        chosen_display = st.multiselect(\"Wähle Artikel\", options=options, default=options if select_all else [])
        chosen = [mapping[d] for d in chosen_display]
    else:
        st.info(\"Keine Artikel verfügbar. Lade die Excel-Datei hoch.\")
        chosen = []

    st.markdown(\"---\")
    st.subheader(\"Vorlagen (Templates)\")
    if \"templates\" not in st.session_state:
        st.session_state[\"templates\"] = []
    with st.expander(\"Vorlage importieren / exportieren\"):
        col1, col2 = st.columns(2)
        with col1:
            tpl_upload = st.file_uploader(\"Vorlage (JSON) importieren\", type=[\"json\"], key=\"tpl_upload\")
            if tpl_upload is not None:
                try:
                    data = json.load(tpl_upload)
                    st.session_state[\"templates\"].append(data)
                    st.success(\"Vorlage importiert.\")
                except Exception as e:
                    st.error(\"Fehler beim Laden der Vorlage.\")
        with col2:
            if st.session_state[\"templates\"]:
                to_export = json.dumps(st.session_state[\"templates\"][-1], ensure_ascii=False, indent=2)
                st.download_button(\"Aktuelle Vorlage herunterladen (JSON)\", data=to_export, file_name=\"template.json\", mime=\"application/json\")
            else:
                st.info(\"Keine Vorlage vorhanden. Erstelle zuerst eine Vorlage aus dem Paket rechts.\")

with right:
    st.title(\"Paket konfigurieren\")
    if df.empty:
        st.info(\"Bitte lade die Excel-Datei hoch oder lege sie als 'Miete.xlsx' neben die App-Datei.\")
        st.stop()

    name_col = df.columns[0]
    price_col = df.columns[-1]
    df[price_col] = pd.to_numeric(df[price_col], errors='coerce').fillna(0.0)

    if \"package\" not in st.session_state:
        st.session_state[\"package\"] = []

    for art in chosen:
        if art and all(not (p[\"Artikel\"] == art) for p in st.session_state[\"package\"]):
            row = df[df[name_col] == art].iloc[0]
            st.session_state[\"package\"].append({
                \"Artikel\": art,
                \"Menge\": 1,
                \"Einzelpreis\": float(row[price_col]),
                \"ZeileNetto\": float(row[price_col]) * 1
            })

    st.subheader(\"Positionen\")
    if not st.session_state[\"package\"]:
        st.info(\"Füge Artikel aus der linken Liste hinzu oder lade eine Vorlage.\")
    else:
        remove_indices = []
        package_lines = st.session_state[\"package\"]
        cols = st.columns([5,2,2,2,1])
        cols[0].markdown(\"**Artikel**\")
        cols[1].markdown(\"**Menge**\")
        cols[2].markdown(\"**Einzelpreis**\")
        cols[3].markdown(\"**Zeile (Netto)**\")
        cols[4].markdown(\"**Aktion**\")
        for i, line in enumerate(package_lines):
            cols = st.columns([5,2,2,2,1])
            cols[0].write(line[\"Artikel\"])
            qty = cols[1].number_input(f\"qty_{i}\", min_value=0, value=int(line.get(\"Menge\",1)), step=1, key=f\"qty_{i}\")
            price = float(line.get(\"Einzelpreis\",0.0))
            line_net = qty * price
            cols[2].write(f\"{price:.2f} €\")
            cols[3].write(f\"{line_net:.2f} €\")
            if cols[4].button(\"Entfernen\", key=f\"remove_{i}\"):
                remove_indices.append(i)
            st.session_state[\"package\"][i][\"Menge\"] = int(qty)
            st.session_state[\"package\"][i][\"ZeileNetto\"] = float(line_net)

        for idx in sorted(remove_indices, reverse=True):
            st.session_state[\"package\"].pop(idx)
        pkg_df = pd.DataFrame(st.session_state[\"package"])

        st.markdown(\"---\")
        st.subheader(\"Zusammenfassung\")
        net_sum = pkg_df[\"ZeileNetto\"].sum() if not pkg_df.empty else 0.0

        col1, col2, col3 = st.columns(3)
        with col1:
            discount_pct = st.number_input(\"Rabatt (%)\", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
        with col2:
            mwst_pct = st.selectbox(\"MwSt (%)\", options=[0,7,19], index=2)
        with col3:
            st.write(\"\")
            if st.button(\"Als Vorlage speichern\"):
                tpl = {\"name\": f\"Vorlage_{datetime.now().strftime('%Y%m%d_%H%M%S')}\", \"package\": st.session_state[\"package\"], \"mwst\": mwst_pct, \"discount\": discount_pct}
                st.session_state[\"templates\"].append(tpl)
                st.success(\"Vorlage gespeichert. Du kannst sie links herunterladen.\")

        discount_amount = net_sum * (discount_pct/100.0)
        net_after_discount = net_sum - discount_amount
        vat_amount = net_after_discount * (mwst_pct/100.0)
        gross_total = net_after_discount + vat_amount

        st.markdown(f\"**Zwischensumme (Netto):** {net_sum:.2f} €\")
        if discount_pct > 0:
            st.markdown(f\"**Rabatt ({discount_pct:.2f}%):** −{discount_amount:.2f} €\")
        st.markdown(f\"**MwSt ({mwst_pct}%):** {vat_amount:.2f} €\")
        st.markdown(f\"**Gesamtbetrag (Brutto):** {gross_total:.2f} €\")

        st.markdown(\"---\")
        def to_excel_bytes(df_export):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine=\"openpyxl\") as writer:
                df_export.to_excel(writer, index=False, sheet_name=\"Paket\")
                writer.save()
            return output.getvalue()

        export_df = pkg_df.copy()
        export_df[\"Rabatt_%\"] = discount_pct
        export_df[\"MwSt_%\"] = mwst_pct
        export_df[\"NettoSumme\"] = net_sum
        export_df[\"NettoNachRabatt\"] = net_after_discount
        export_df[\"MwStBetrag\"] = vat_amount
        export_df[\"BruttoSumme\"] = gross_total
        filename = f\"Paket_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx\"
        st.download_button(\"📥 Paket als Excel exportieren\", data=to_excel_bytes(export_df), file_name=filename, mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\")

        offer_lines = [f\"Angebot — Paket vom {datetime.now().strftime('%Y-%m-%d')}\"]
        offer_lines.append(\"Positionen:\")
        for _, r in export_df.iterrows():
            offer_lines.append(f\"- {int(r['Menge'])} × {r['Artikel']} @ {r['Einzelpreis']:.2f} € = {r['ZeileNetto']:.2f} €\")
        offer_lines.append(f\"\\nZwischensumme: {net_sum:.2f} €\")
        if discount_pct > 0:
            offer_lines.append(f\"Rabatt: {discount_amount:.2f} € ({discount_pct:.2f}%)\")
        offer_lines.append(f\"MwSt: {vat_amount:.2f} € ({mwst_pct}%)\")
        offer_lines.append(f\"Gesamt (Brutto): {gross_total:.2f} €\")
        offer_text = \"\\n\".join(offer_lines)

        st.subheader(\"Angebotstext\")
        st.code(offer_text, language=\"text\")
        st.download_button(\"📄 Angebotstext als .txt herunterladen\", data=offer_text, file_name=filename.replace('.xlsx', '.txt'), mime=\"text/plain\")

    if st.session_state.get(\"templates\"):
        st.markdown(\"---\")
        st.subheader(\"Gespeicherte Vorlagen\")
        for ti, tpl in enumerate(st.session_state[\"templates\"][::-1]):
            cols = st.columns([6,2,2])
            cols[0].write(tpl.get(\"name\", f\"Vorlage {ti}\"))
            if cols[1].button(\"Laden\", key=f\"load_tpl_{ti}\"):
                st.session_state[\"package\"] = tpl[\"package\"].copy()
                st.success(\"Vorlage geladen in das aktuelle Paket.\")
            if cols[2].button(\"Herunterladen\", key=f\"dl_tpl_{ti}\"):
                st.download_button(\"Download\", data=json.dumps(tpl, ensure_ascii=False, indent=2), file_name=f\"{tpl.get('name','template')}.json\", mime=\"application/json\")
