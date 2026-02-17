import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.worksheet.worksheet import Worksheet

# --- Configuration de la page ---
st.set_page_config(page_title="Rapprochement Facture vs BL", page_icon="⚖️", layout="wide")

# --- Constantes des Colonnes ---
COL_BL_DATE = 'Date Document'
COL_BL_NUM = 'N° de pièce'
COL_BL_REF_ART = 'Référence Article'
COL_BL_REF_FOURN = 'AF_RefFourniss' # Clé de jointure BL
COL_BL_DES = 'Désignation Article'
COL_BL_PU = 'Prix Unitaire HT'
COL_BL_QTE = 'Qté Livrées'
COL_BL_REM = 'Pourcentage Remise'
COL_BL_MNT = 'Montant HT Net'

COL_FAC_REF_FOURN = 'Af_RefFourniss' # Clé de jointure Facture
COL_FAC_PU = 'Prix unitaire'
COL_FAC_REM = 'Remise'
COL_FAC_QTE = 'Quantité'
COL_FAC_MNT = 'Montant HT'

# Colonnes Calculées
COL_ECART_QTE = 'Écart Qté'
COL_ECART_PRIX = 'Écart Prix U.'
COL_ECART_REM = 'Écart Remise (%)'
COL_ECART_MNT = 'Écart Montant HT'
COL_STATUT = 'Statut'

# --- Initialisation de l'État de Session ---
defaults = {
    'data_loaded': False,
    'df_bl_raw': pd.DataFrame(),
    'df_fac_raw': pd.DataFrame(),
    'df_resultat': pd.DataFrame(),
    'manual_map': {}
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- Fonctions Utilitaires ---
def clean_currency(series):
    return pd.to_numeric(
        series.astype(str).str.replace('€', '').str.replace(' ', '').str.replace(',', '.'), 
        errors='coerce'
    ).fillna(0.0)

def clean_text_id(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

def nettoyer_ref(series):
    return series.astype(str).str.strip().str.upper()

# --- Export Excel ---
def apply_excel_styling_rapprochement(ws: Worksheet, df_columns: list):
    """Applique le style standard du tableau de rapprochement."""
    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    error_fill = PatternFill(start_color="FFC7CE", fill_type="solid")
    error_font = Font(color="9C0006")
    ok_font = Font(color="006100")
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    euro_format = '#,##0.00 €'
    qty_format = '0.00'

    cols_euro = [COL_BL_PU, COL_BL_MNT, COL_FAC_PU, COL_FAC_MNT, COL_ECART_PRIX, COL_ECART_MNT]
    cols_qty = [COL_BL_QTE, COL_FAC_QTE, COL_ECART_QTE, COL_BL_REM, COL_FAC_REM, COL_ECART_REM]
    cols_text = [COL_BL_NUM]

    for row_idx, row in enumerate(ws.iter_rows(), 1):
        for cell in row:
            col_name = df_columns[cell.column - 1]
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal='center')
            else:
                if col_name in cols_euro: cell.number_format = euro_format
                elif col_name in cols_qty: cell.number_format = qty_format
                elif col_name in cols_text: 
                    cell.number_format = '@'
                    cell.alignment = Alignment(horizontal='left')
                
                if col_name in [COL_ECART_QTE, COL_ECART_PRIX, COL_ECART_MNT, COL_ECART_REM]:
                    val = cell.value if isinstance(cell.value, (int, float)) else 0
                    if abs(val) > 0.01:
                        cell.fill = error_fill; cell.font = error_font
                    else:
                        cell.font = ok_font

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = min((max_length + 2) * 1.2, 50)

def style_recap_sheet(ws: Worksheet):
    """Applique un style spécifique simple pour l'onglet Récapitulatif."""
    header_fill = PatternFill(start_color="D9D9D9", fill_type="solid")
    bold_font = Font(bold=True)
    euro_format = '#,##0.00 €'
    
    # Largeur colonnes
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 25

    for row in ws.iter_rows():
        cell_key = row[0]
        cell_val = row[1]
        
        cell_key.font = bold_font
        
        # Format monétaire pour les 3 premières lignes (Montants)
        if cell_key.row <= 3: 
             cell_val.number_format = euro_format
             cell_val.font = bold_font

def generer_excel(df):
    output = io.BytesIO()
    
    # --- Calcul des indicateurs pour le Récapitulatif ---
    tot_fac = df[COL_FAC_MNT].sum()
    tot_bl = df[COL_BL_MNT].sum()
    ecart_glob = tot_fac - tot_bl
    
    # Filtres
    mask_prix = abs(df[COL_ECART_PRIX]) > 0.01
    mask_rem = abs(df[COL_ECART_REM]) > 0.01
    mask_manque_bl = df[COL_STATUT] == 'Manque BL' # Facturé mais pas livré
    mask_non_fac = df[COL_STATUT] == 'Non Facturé' # Livré mais pas facturé
    mask_ecarts_generaux = (df[COL_STATUT] == 'Écart Prix/Remise/Qté') # Lignes avec écarts valeurs
    
    nb_ecarts_constates = mask_ecarts_generaux.sum()
    nb_ecarts_prix = mask_prix.sum()
    nb_ecarts_rem = mask_rem.sum()
    nb_abs_bl = mask_manque_bl.sum()
    nb_abs_fac = mask_non_fac.sum()

    # Création du DataFrame Récapitulatif
    data_recap = [
        ("Montant total de la facture", tot_fac),
        ("Montant total des BL", tot_bl),
        ("Écart Global", ecart_glob),
        ("", ""), # Ligne vide
        ("Nombre de produits avec écarts constatés", nb_ecarts_constates),
        ("Nombre d'écarts de prix", nb_ecarts_prix),
        ("Nombre d'écarts de remises", nb_ecarts_rem),
        ("Nombre de produits facturés non trouvés dans les BL", nb_abs_bl),
        ("Nombre de produits sur les BL non trouvés dans la facture", nb_abs_fac)
    ]
    df_recap = pd.DataFrame(data_recap, columns=["Indicateur", "Valeur"])

    # --- Écriture Excel Multi-onglets ---
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # 1. Onglet Récapitulatif
        df_recap.to_excel(writer, sheet_name='Récapitulatif', index=False, header=False)
        style_recap_sheet(writer.sheets['Récapitulatif'])
        
        # 2. Onglet Global (Tout)
        df.to_excel(writer, sheet_name='Rapprochement Global', index=False)
        apply_excel_styling_rapprochement(writer.sheets['Rapprochement Global'], df.columns.tolist())
        
        # 3. Onglet Écarts de Prix
        if nb_ecarts_prix > 0:
            df_p = df[mask_prix].copy()
            df_p.to_excel(writer, sheet_name='Ecarts Prix', index=False)
            apply_excel_styling_rapprochement(writer.sheets['Ecarts Prix'], df.columns.tolist())
            
        # 4. Onglet Écarts de Remise
        if nb_ecarts_rem > 0:
            df_r = df[mask_rem].copy()
            df_r.to_excel(writer, sheet_name='Ecarts Remise', index=False)
            apply_excel_styling_rapprochement(writer.sheets['Ecarts Remise'], df.columns.tolist())
            
        # 5. Onglet Facturés mais Absents BL
        if nb_abs_bl > 0:
            df_abl = df[mask_manque_bl].copy()
            df_abl.to_excel(writer, sheet_name='Absents BL', index=False)
            apply_excel_styling_rapprochement(writer.sheets['Absents BL'], df.columns.tolist())

        # 6. Onglet BL mais Non Facturés
        if nb_abs_fac > 0:
            df_anf = df[mask_non_fac].copy()
            df_anf.to_excel(writer, sheet_name='Non Facturés', index=False)
            apply_excel_styling_rapprochement(writer.sheets['Non Facturés'], df.columns.tolist())

    output.seek(0)
    return output

# --- Interface Principale ---
st.title("⚖️ Rapprochement Facture Fournisseur vs BL")

with st.sidebar:
    st.header("1. Importation")
    if st.button("🔄️ Réinitialiser tout"):
        st.session_state.clear()
        st.rerun()
        
    mode = st.radio("Source des données", ["2 Fichiers distincts", "1 Fichier (onglets)"])
    file_bl, file_fac, sheet_bl, sheet_fac = None, None, 0, 0
    
    if mode == "2 Fichiers distincts":
        file_fac = st.file_uploader("Fichier FACTURE", type=["xlsx"])
        file_bl = st.file_uploader("Fichier BL", type=["xlsx"])
    else:
        file_unique = st.file_uploader("Fichier Excel complet", type=["xlsx"])
        if file_unique:
            xl = pd.ExcelFile(file_unique)
            sheet_fac = st.selectbox("Onglet FACTURE", xl.sheet_names, index=0)
            sheet_bl = st.selectbox("Onglet BL", xl.sheet_names, index=1 if len(xl.sheet_names)>1 else 0)
            file_fac = file_unique; file_bl = file_unique

    if file_fac and file_bl:
        if st.button("Charger les données", type="primary"):
            try:
                df_fac = pd.read_excel(file_fac, sheet_name=sheet_fac)
                df_bl = pd.read_excel(file_bl, sheet_name=sheet_bl, header=2)
                
                if COL_BL_REF_FOURN in df_bl.columns and COL_FAC_REF_FOURN in df_fac.columns:
                    # BL Clean
                    for col in [COL_BL_PU, COL_BL_QTE, COL_BL_REM, COL_BL_MNT]: 
                        if col in df_bl.columns: df_bl[col] = clean_currency(df_bl[col])
                    if COL_BL_NUM in df_bl.columns:
                        df_bl[COL_BL_NUM] = clean_text_id(df_bl[COL_BL_NUM])
                    
                    # Facture Clean
                    for col in [COL_FAC_PU, COL_FAC_QTE, COL_FAC_REM, COL_FAC_MNT]:
                        if col in df_fac.columns: df_fac[col] = clean_currency(df_fac[col])
                    
                    df_bl.insert(0, "Sélectionner", True)
                    st.session_state.df_bl_raw = df_bl
                    st.session_state.df_fac_raw = df_fac
                    st.session_state.manual_map = {}
                    st.session_state.data_loaded = True
                    st.rerun()
                else:
                    st.error(f"Colonnes clés manquantes : {COL_BL_REF_FOURN} ou {COL_FAC_REF_FOURN}")
            except Exception as e:
                st.error(f"Erreur lecture : {e}")

if st.session_state.data_loaded:
    tab1, tab2, tab3 = st.tabs(["✅ Sélection BL", "🔗 Rapprochement Manuel", "📊 Résultats & Export"])
    
    # ---------------------------------------------------------
    # TAB 1 : SÉLECTION DES BL
    # ---------------------------------------------------------
    with tab1:
        st.info("Décochez les lignes de BL à ignorer.")
        cols_view = ["Sélectionner", COL_BL_NUM, COL_BL_REF_FOURN, COL_BL_DES, COL_BL_PU, COL_BL_REM, COL_BL_QTE]
        cols_exist = [c for c in cols_view if c in st.session_state.df_bl_raw.columns]
        
        edited_bl = st.data_editor(
            st.session_state.df_bl_raw[cols_exist],
            column_config={"Sélectionner": st.column_config.CheckboxColumn(default=True)},
            disabled=[c for c in cols_exist if c != "Sélectionner"],
            hide_index=True, use_container_width=True, height=400, key="editor_bl"
        )
    
    bl_selected_indices = edited_bl[edited_bl["Sélectionner"] == True].index
    df_bl_active = st.session_state.df_bl_raw.loc[bl_selected_indices].copy()
    
    # ---------------------------------------------------------
    # TAB 2 : RAPPROCHEMENT MANUEL
    # ---------------------------------------------------------
    with tab2:
        st.markdown("#### 🔗 Lier des références orphelines")
        
        df_bl_mapped = df_bl_active.copy()
        df_bl_mapped['Ref_Join'] = nettoyer_ref(df_bl_mapped[COL_BL_REF_FOURN])
        for ref_bl_orig, ref_fac_target in st.session_state.manual_map.items():
            mask = df_bl_mapped['Ref_Join'] == nettoyer_ref(pd.Series([ref_bl_orig]))[0]
            df_bl_mapped.loc[mask, 'Ref_Join'] = nettoyer_ref(pd.Series([ref_fac_target]))[0]
            
        refs_bl = set(df_bl_mapped['Ref_Join'].unique())
        df_fac_clean = st.session_state.df_fac_raw.copy()
        df_fac_clean['Ref_Join'] = nettoyer_ref(df_fac_clean[COL_FAC_REF_FOURN])
        refs_fac = set(df_fac_clean['Ref_Join'].unique())
        
        orphans_fac = sorted(list(refs_fac - refs_bl))
        orphans_bl = sorted(list(refs_bl - refs_fac))
        
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            sel_fac = st.selectbox("1. Réf Facture orpheline", orphans_fac, key="sel_orph_fac") if orphans_fac else None
        with c2:
            if orphans_bl:
                 def get_des(ref):
                     rows = df_bl_active[nettoyer_ref(df_bl_active[COL_BL_REF_FOURN]) == ref]
                     return str(rows.iloc[0].get(COL_BL_DES, ''))[:30] if not rows.empty else ""
                 opts_bl = [f"{r} | {get_des(r)}" for r in orphans_bl]
                 sel_bl_display = st.selectbox("2. Réf BL correspondante", opts_bl, key="sel_orph_bl")
                 sel_bl = sel_bl_display.split(" | ")[0] if sel_bl_display else None
            else: sel_bl = None
                
        with c3:
            st.write("Action")
            if st.button("➕ Créer le lien") and sel_fac and sel_bl:
                st.session_state.manual_map[sel_bl] = sel_fac
                st.success(f"Lien créé : {sel_bl} ➡️ {sel_fac}")
                st.rerun()

        if st.session_state.manual_map:
            st.divider()
            st.write("**Liens actifs :**")
            to_del = None
            for k, v in st.session_state.manual_map.items():
                c_a, c_b = st.columns([4, 1])
                c_a.text(f"BL [{k}] fusionné vers [{v}]")
                if c_b.button("🗑️", key=f"del_{k}"): to_del = k
            if to_del: del st.session_state.manual_map[to_del]; st.rerun()

    # ---------------------------------------------------------
    # TAB 3 : RÉSULTATS
    # ---------------------------------------------------------
    with tab3:
        if st.button("🚀 Lancer le calcul", type="primary"):
            
            # --- 1. PRÉPARATION BL ---
            df_bl_calc = df_bl_active.copy()
            df_bl_calc['_key_orig'] = nettoyer_ref(df_bl_calc[COL_BL_REF_FOURN])
            
            # Application du mapping
            df_bl_calc['_key'] = df_bl_calc['_key_orig']
            for ref_bl, ref_fac in st.session_state.manual_map.items():
                clean_r_bl = nettoyer_ref(pd.Series([ref_bl]))[0]
                clean_r_fac = nettoyer_ref(pd.Series([ref_fac]))[0]
                df_bl_calc.loc[df_bl_calc['_key_orig'] == clean_r_bl, '_key'] = clean_r_fac
            
            # GroupBy BL
            grp_bl = df_bl_calc.groupby('_key').agg({
                COL_BL_REF_FOURN: 'first',
                COL_BL_QTE: 'sum',
                COL_BL_MNT: 'sum',
                COL_BL_PU: 'first',
                COL_BL_REM: 'first',
                COL_BL_DES: 'first',
                COL_BL_NUM: lambda x: ', '.join(sorted(list(set(x.astype(str)))))
            }).reset_index()

            # --- 2. PRÉPARATION FACTURE ---
            df_fac_calc = st.session_state.df_fac_raw.copy()
            df_fac_calc['_key'] = nettoyer_ref(df_fac_calc[COL_FAC_REF_FOURN])
            
            grp_fac = df_fac_calc.groupby('_key').agg({
                COL_FAC_REF_FOURN: 'first',
                COL_FAC_QTE: 'sum',
                COL_FAC_MNT: 'sum',
                COL_FAC_PU: 'first',
                COL_FAC_REM: 'first'
            }).reset_index()
            
            # --- 3. FUSION ---
            merged = pd.merge(grp_fac, grp_bl, on='_key', how='outer', suffixes=('_FAC', '_BL'), indicator=True)
            
            # --- 4. CALCULS ---
            cols_num = [COL_FAC_QTE, COL_BL_QTE, COL_FAC_MNT, COL_BL_MNT, COL_FAC_PU, COL_BL_PU, COL_FAC_REM, COL_BL_REM]
            for c in cols_num: merged[c] = merged[c].fillna(0)
            
            merged[COL_ECART_QTE] = merged[COL_FAC_QTE] - merged[COL_BL_QTE]
            merged[COL_ECART_MNT] = merged[COL_FAC_MNT] - merged[COL_BL_MNT]
            merged[COL_ECART_PRIX] = merged[COL_FAC_PU] - merged[COL_BL_PU]
            merged[COL_ECART_REM] = merged[COL_FAC_REM] - merged[COL_BL_REM]
            
            conditions = [
                (merged['_merge'] == 'left_only'),
                (merged['_merge'] == 'right_only'),
                (abs(merged[COL_ECART_MNT]) > 0.05) | (abs(merged[COL_ECART_REM]) > 0.01)
            ]
            choices = ['Manque BL', 'Non Facturé', 'Écart Prix/Remise/Qté']
            merged[COL_STATUT] = np.select(conditions, choices, default='OK')
            
            merged['Référence'] = merged[COL_FAC_REF_FOURN].fillna(merged[COL_BL_REF_FOURN] + " (BL)")
            
            final_cols = [
                'Référence', COL_BL_DES, COL_BL_NUM,
                COL_FAC_QTE, COL_FAC_PU, COL_FAC_REM, COL_FAC_MNT,
                COL_BL_QTE, COL_BL_PU, COL_BL_REM, COL_BL_MNT,
                COL_ECART_QTE, COL_ECART_PRIX, COL_ECART_REM, COL_ECART_MNT,
                COL_STATUT
            ]
            st.session_state.df_resultat = merged[final_cols].copy()

        if not st.session_state.df_resultat.empty:
            df_res = st.session_state.df_resultat
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Facture", f"{df_res[COL_FAC_MNT].sum():,.2f} €")
            c2.metric("Total BL", f"{df_res[COL_BL_MNT].sum():,.2f} €")
            c3.metric("Écart Global", f"{(df_res[COL_FAC_MNT].sum() - df_res[COL_BL_MNT].sum()):,.2f} €", delta_color="inverse")
            
            def highlight_errors(val):
                return 'background-color: #ffcccc; color: #990000' if isinstance(val, (int, float)) and abs(val) > 0.01 else ''
            
            st.dataframe(
                df_res.style.map(highlight_errors, subset=[COL_ECART_QTE, COL_ECART_PRIX, COL_ECART_REM, COL_ECART_MNT])
                            .format({COL_BL_REM: "{:.2f}", COL_FAC_REM: "{:.2f}", COL_ECART_REM: "{:.2f}"}),
                use_container_width=True, height=500
            )
            
            fname = st.text_input("Nom du fichier", "Resultat_Rapprochement")
            st.download_button("📥 Télécharger Excel complet", data=generer_excel(df_res), file_name=f"{fname}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
