# Accueil.py (Script Principal / Page d'Accueil - VERSION MISE À JOUR)

import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import calendar

# Importer les fonctions de helpers.py
try:
    from helpers import (
        safe_read_excel, parse_week_column_to_date
    )
except ImportError:
    st.error("Le fichier 'helpers.py' est introuvable ou les fonctions requises sont manquantes. Certaines fonctionnalités pourraient ne pas marcher.")
    # Définir des fonctions de remplacement pour éviter les crashs
    def safe_read_excel(*args, **kwargs): return None
    def parse_week_column_to_date(*args, **kwargs): return None

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Définition de l'état de session par défaut ---
def get_master_default_session_state():
    return {
        'df_full': None, 'min_order_dict': {}, 'df_initial_filtered': pd.DataFrame(),
        'all_available_semaine_columns': [], 'unique_suppliers_list': [],
        'df_product_events': pd.DataFrame(), 'df_suivi_commandes_all': pd.DataFrame(),
        'supplier_evaluation_data': None, 'global_stock_target_config': 3200000.0,
        'commande_result_df': None, 'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [], 'commande_params_calculated_for': {},
        'ai_commande_result_df': None, 'ai_commande_total_amount': 0.0,
        'ai_commande_params_calculated_for': {}, 'ai_forecast_weeks_val': 4,
        'ai_min_order_val': 0.0, 'ai_ignored_orders_df': pd.DataFrame(),
        'ai_excluded_suppliers_stock_target': [],
        'rotation_result_df': None, 'rotation_analysis_period_label': "12 dernières semaines",
        'rotation_threshold_value': 1.0, 'show_all_rotation_data': True,
        'rotation_params_calculated_for': {}, 'rotation_ia_result_df': None,
        'rotation_ia_params_calculated_for': {}, 'rotation_ia_projection_weeks': 4,
        'forecast_result_df': None, 'forecast_grand_total_amount': 0.0,
        'forecast_simulation_params_calculated_for': {},
        'forecast_selected_months_ui': list(calendar.month_name)[1:],
        'forecast_sim_type_radio_index': 0, 'forecast_progression_percentage_ui': 5.0,
        'forecast_target_amount_ui': 10000.0,

        # --- CLÉS POUR L'APPLICATION RFA ---
        'rfa_result_df': None,
        'rfa_params_calculated_for': {},
        'rfa_supplier_selection_ui': [],
        'rfa_seuil_a_ui': 80.0,
        'rfa_seuil_b_ui': 95.0,

        # --- CLÉS POUR L'APPLICATION 7_STOCKS ---
        'stocks_result_df': None,
        'stocks_prophet_forecast_df': None,
        'stocks_params_calculated_for': {},
        'stocks_prophet_horizon_ui': 90,
        'stocks_eoy_safety_stock_units_ui': 500,
        # ----------------------------------------------------

        'uploaded_file_name': None,
        'session_state_initialized': False,
        'debug_last_file_processed_name': None,
        'debug_df_full_was_set': False,
        'ai_debug_logs_display_trigger': False,
        'ai_debug_logs_expanded_state': False
    }

if not st.session_state.get('session_state_initialized', False):
    default_state_values = get_master_default_session_state()
    for key, default_value in default_state_values.items():
        if key not in st.session_state:
            st.session_state[key] = default_value
    st.session_state.session_state_initialized = True
    logging.info("Master session state initialized from Accueil.py.")

st.set_page_config(
    page_title="Tableau de Bord Logistique", layout="wide", initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'mailto:votre.email@example.com',
        'Report a bug': "mailto:votre.email@example.com",
        'About': "# Application d'aide à la décision logistique."
    }
)

# --- CONTENU DE LA PAGE D'ACCUEIL ---
st.title("📊 Tableau de Bord Logistique Global")
st.markdown("Chargez votre fichier de travail Excel principal ici pour activer les outils d'analyse disponibles dans les autres pages (via la barre latérale).")
st.markdown("---")

with st.expander("ℹ️ Informations de Débogage Session State (Accueil)", expanded=False):
    st.write(f"Nom du fichier uploadé (session): `{st.session_state.get('uploaded_file_name')}`")
    st.write(f"df_full est None (session): `{st.session_state.get('df_full') is None}`")
    st.write(f"Fichier traité (debug flag): `{st.session_state.get('debug_last_file_processed_name')}`")
    st.write(f"df_full a été assigné (debug flag): `{st.session_state.get('debug_df_full_was_set')}`")
    df_events_debug = st.session_state.get('df_product_events')
    if df_events_debug is not None:
        st.write(f"df_product_events lignes (session): `{len(df_events_debug)}`")

uploaded_file = st.file_uploader(
    "📁 Charger votre fichier Excel principal", type=["xlsx", "xls"],
    key="master_file_uploader_accueil",
    help="Le fichier doit contenir les onglets 'Tableau final', 'Special', 'Minimum de commande', 'Suivi commandes'."
)

if uploaded_file:
    session_uploaded_name = st.session_state.get('uploaded_file_name')
    session_df_full_is_none = st.session_state.get('df_full') is None
    
    trigger_reload = False
    if session_uploaded_name != uploaded_file.name: trigger_reload = True
    elif session_df_full_is_none: trigger_reload = True
    elif session_uploaded_name == uploaded_file.name and session_df_full_is_none and \
         (st.session_state.get('debug_last_file_processed_name') != uploaded_file.name or \
          not st.session_state.get('debug_df_full_was_set', False) ): trigger_reload = True

    if trigger_reload:
        st.warning(f"Déclenchement du traitement du fichier : {uploaded_file.name}")
        
        keys_to_clear_on_new_file = list(get_master_default_session_state().keys())
        preserved_keys = ['session_state_initialized',
                          'ai_debug_logs_display_trigger', 'ai_debug_logs_expanded_state']

        st.session_state.uploaded_file_name = uploaded_file.name

        for pk in preserved_keys:
            if pk in keys_to_clear_on_new_file: keys_to_clear_on_new_file.remove(pk)
        if 'uploaded_file_name' in keys_to_clear_on_new_file:
             keys_to_clear_on_new_file.remove('uploaded_file_name')

        default_state_after_clear = get_master_default_session_state()
        for key_to_reset in keys_to_clear_on_new_file:
            st.session_state[key_to_reset] = default_state_after_clear[key_to_reset]
        
        st.session_state.df_full = None 
        st.session_state.df_initial_filtered = pd.DataFrame()
        st.session_state.df_product_events = pd.DataFrame()
        st.session_state.df_suivi_commandes_all = pd.DataFrame()
        st.session_state.min_order_dict = {}
        st.session_state.all_available_semaine_columns = []
        st.session_state.unique_suppliers_list = []
        st.session_state.ai_ignored_orders_df = pd.DataFrame()
        if 'ai_debug_logs' in st.session_state: st.session_state.ai_debug_logs = []

        st.session_state.debug_last_file_processed_name = None 
        st.session_state.debug_df_full_was_set = False      

        st.info("État de session pertinent réinitialisé pour le nouveau fichier.")
        logging.info(f"Traitement du fichier: {uploaded_file.name}.")
        progress_bar = st.progress(0, text="Chargement du fichier...")
            
        try:
            excel_io_buf = io.BytesIO(uploaded_file.getvalue())
            
            progress_bar.progress(10, text="Lecture 'Tableau final'...")
            df_full_read = safe_read_excel(excel_io_buf, sheet_name="Tableau final", header=7)
            if df_full_read is None or df_full_read.empty: 
                st.error("❌ L'onglet 'Tableau final' n'a pas pu être lu ou est vide."); st.stop()
            
            req_tf_cols_check = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article", "Date Création Article"]
            missing_tf_check = [c for c in req_tf_cols_check if c not in df_full_read.columns]
            if missing_tf_check: 
                st.error(f"❌ Colonnes manquantes dans 'Tableau final': {', '.join(missing_tf_check)}."); st.stop()

            df_full_read["Stock"] = pd.to_numeric(df_full_read["Stock"], errors='coerce').fillna(0)
            df_full_read["Tarif d'achat"] = pd.to_numeric(df_full_read["Tarif d'achat"], errors='coerce').fillna(0)
            df_full_read["Conditionnement"] = pd.to_numeric(df_full_read["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: int(x) if x > 0 else 1)
            if "Date Création Article" in df_full_read.columns:
                df_full_read["Date Création Article"] = pd.to_datetime(df_full_read["Date Création Article"], errors='coerce')
            for str_c_tf in ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]:
                if str_c_tf in df_full_read.columns: df_full_read[str_c_tf] = df_full_read[str_c_tf].astype(str).str.strip().replace('nan', '')
            
            st.session_state.df_full = df_full_read
            st.info(f"'Tableau final' lu et traité ({len(st.session_state.df_full)} lignes).")
            progress_bar.progress(25, text="'Tableau final' traité.")

            progress_bar.progress(30, text="Lecture 'Min cmd'...")
            excel_io_buf.seek(0)
            df_min_c_read = safe_read_excel(excel_io_buf, sheet_name="Minimum de commande")
            min_o_dict_temp_read = {}
            if df_min_c_read is not None and not df_min_c_read.empty:
                s_col_min, m_col_min = "Fournisseur", "Minimum de Commande"
                if s_col_min in df_min_c_read.columns and m_col_min in df_min_c_read.columns:
                    try:
                        df_min_c_read[s_col_min] = df_min_c_read[s_col_min].astype(str).str.strip().replace('nan', '')
                        df_min_c_read[m_col_min] = pd.to_numeric(df_min_c_read[m_col_min], errors='coerce').astype(float)
                        min_o_dict_temp_read = df_min_c_read.dropna(subset=[s_col_min, m_col_min]).set_index(s_col_min)[m_col_min].to_dict()
                    except Exception as e_min_proc: st.warning(f"⚠️ Erreur traitement 'Min cmd': {e_min_proc}")
            st.session_state.min_order_dict = min_o_dict_temp_read
            progress_bar.progress(40, text=f"'Min cmd' lu ({len(st.session_state.min_order_dict)} entrées).")

            progress_bar.progress(45, text="Lecture 'Suivi cmds'...")
            excel_io_buf.seek(0)
            df_suivi_read = safe_read_excel(excel_io_buf, sheet_name="Suivi commandes", header=4)
            if df_suivi_read is not None and not df_suivi_read.empty:
                req_s_cols_check_suivi = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Intitulé Fournisseur"]
                missing_s_check_suivi = [c for c in req_s_cols_check_suivi if c not in df_suivi_read.columns]
                if not missing_s_check_suivi:
                    df_suivi_read.rename(columns={"Intitulé Fournisseur": "Fournisseur"}, inplace=True)
                    for col_s in ["Fournisseur", "AF_RefFourniss", "Désignation Article", "N° de pièce"]:
                        if col_s in df_suivi_read.columns: df_suivi_read[col_s] = df_suivi_read[col_s].astype(str).str.strip().replace('nan','')
                    if "Qté Commandées" in df_suivi_read.columns: df_suivi_read["Qté Commandées"] = pd.to_numeric(df_suivi_read["Qté Commandées"], errors='coerce').fillna(0)
                    if "Date Pièce BC" in df_suivi_read.columns:
                        df_suivi_read["Date Pièce BC"] = pd.to_datetime(df_suivi_read["Date Pièce BC"], errors='coerce')
                    df_suivi_read.dropna(how='all', inplace=True)
                    st.session_state.df_suivi_commandes_all = df_suivi_read
                else: st.warning(f"⚠️ Cols manquantes dans 'Suivi commandes': {', '.join(missing_s_check_suivi)}. Onglet ignoré.")
            elif df_suivi_read is None: st.info("ℹ️ Onglet 'Suivi commandes' non trouvé ou erreur de lecture.")
            else: st.info("ℹ️ Onglet 'Suivi commandes' est vide.")
            progress_bar.progress(55, text=f"'Suivi cmds' lu ({len(st.session_state.df_suivi_commandes_all)} lignes).")
            
            progress_bar.progress(60, text="Lecture et traitement 'Special' (Ajustements)...")
            excel_io_buf.seek(0) 
            df_events_temp = safe_read_excel(excel_io_buf, sheet_name="Special") 
            
            if df_events_temp is None: st.info("ℹ️ Onglet 'Special' (ajustements) non trouvé ou erreur de lecture.")
            elif df_events_temp.empty: st.info("ℹ️ Onglet 'Special' (ajustements) est vide.")
            else:
                st.info(f"Onglet 'Special' lu ({len(df_events_temp)} lignes brutes). Traitement...")
                df_events_processed_temp = df_events_temp.copy() 
                required_cols_events = ['Référence Article', 'TypeAjustement', 'DateDebut', 'ModeleImpact']
                missing_required_ev = [col for col in required_cols_events if col not in df_events_processed_temp.columns]
                
                if missing_required_ev:
                    st.warning(f"⚠️ Onglet 'Special': Colonnes requises manquantes: {', '.join(missing_required_ev)}. L'onglet sera ignoré ou traité partiellement.")
                else:
                    string_cols_ev = ['Référence Article', 'TypeAjustement', 'ModeleImpact', 'UniteAjustement']
                    for col_ev in string_cols_ev:
                        if col_ev in df_events_processed_temp.columns:
                            df_events_processed_temp[col_ev] = df_events_processed_temp[col_ev].astype(str).str.strip().replace('nan', '')
                            if col_ev in ['ModeleImpact', 'UniteAjustement', 'TypeAjustement']:
                                df_events_processed_temp[col_ev] = df_events_processed_temp[col_ev].str.lower()
                        elif col_ev == 'UniteAjustement':
                             df_events_processed_temp[col_ev] = '' 
                    
                    if 'DateDebut' in df_events_processed_temp.columns: df_events_processed_temp['DateDebut'] = pd.to_datetime(df_events_processed_temp['DateDebut'], errors='coerce')
                    if 'DateFin' in df_events_processed_temp.columns: df_events_processed_temp['DateFin'] = pd.to_datetime(df_events_processed_temp['DateFin'], errors='coerce')
                    
                    if 'ValeurAjustement' in df_events_processed_temp.columns: df_events_processed_temp['ValeurAjustement'] = pd.to_numeric(df_events_processed_temp['ValeurAjustement'], errors='coerce')
                    else: df_events_processed_temp['ValeurAjustement'] = np.nan
                    
                    df_events_processed_temp.dropna(subset=['Référence Article', 'TypeAjustement', 'DateDebut', 'ModeleImpact'], inplace=True)
                    df_events_processed_temp = df_events_processed_temp[
                        (df_events_processed_temp['Référence Article'].astype(str).str.strip() != '') &
                        (df_events_processed_temp['TypeAjustement'].astype(str).str.strip() != '') &
                        (df_events_processed_temp['ModeleImpact'].astype(str).str.strip() != '')
                    ].copy()
                st.session_state.df_product_events = df_events_processed_temp 
            progress_bar.progress(70, text=f"'Special' traité ({len(st.session_state.df_product_events)} événements valides).")

            progress_bar.progress(75, text="Filtrage initial des produits...")
            df_fs_loc_filter = st.session_state.get('df_full')
            if df_fs_loc_filter is not None and not df_fs_loc_filter.empty:
                if "Fournisseur" in df_fs_loc_filter.columns:
                    base_filter_supplier = ((df_fs_loc_filter["Fournisseur"].astype(str).str.strip() != "") & \
                                            (df_fs_loc_filter["Fournisseur"].astype(str).str.strip().str.lower() != "#filter") & \
                                            (df_fs_loc_filter["Fournisseur"].notna()))
                    if "AF_RefFourniss" in df_fs_loc_filter.columns:
                        filter_conditions = base_filter_supplier & \
                                            (df_fs_loc_filter["AF_RefFourniss"].astype(str).str.strip() != "") & \
                                            (df_fs_loc_filter["AF_RefFourniss"].notna())
                        st.session_state.df_initial_filtered = df_fs_loc_filter[filter_conditions].copy()
                    else:
                        st.warning("Colonne 'AF_RefFourniss' non trouvée, filtrage basé sur 'Fournisseur'.")
                        st.session_state.df_initial_filtered = df_fs_loc_filter[base_filter_supplier].copy()
                else: st.warning("Colonne 'Fournisseur' non trouvée. Impossible d'effectuer le filtrage initial.")
            else: st.warning("'Tableau final' vide ou non chargé. Pas de filtrage initial.")
            st.info(f"Filtrage initial effectué: {len(st.session_state.df_initial_filtered)} articles retenus.")

            progress_bar.progress(80, text="Détection des colonnes de ventes (semaines)...")
            st.session_state.all_available_semaine_columns = [] 
            if df_fs_loc_filter is not None and not df_fs_loc_filter.empty and "Désignation Article" in df_fs_loc_filter.columns:
                try:
                    first_week_col_idx_approx = df_fs_loc_filter.columns.get_loc("Désignation Article") + 1
                    potential_sem_cols_read = []
                    
                    if len(df_fs_loc_filter.columns) > first_week_col_idx_approx:
                        candidate_cols_sem = df_fs_loc_filter.columns[first_week_col_idx_approx:].tolist()
                        
                        known_non_week_cols_set = set([
                            "Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", 
                            "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                            "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article",
                            "Désignation Article", "Date Création Article", "Qte Cmdée",
                            "Vts N-1 Total (calc)", "Vts 12 N-1 Sim (calc)", "Vts 12 Dern. (calc)",
                            "Tarif Ach.", "Total Cmd (€)", "Stock Terme", "Qté Cmdée (IA)",
                            "Forecast Ventes (IA)", "Total Cmd (€) (IA)", "Stock Terme (IA)",
                            "WoS_Calculated_Supplier", "SRM_Qty", "Unités Vendues (Période)",
                            "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)",
                            "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "COGS (Période)",
                            "Valeur Stock Actuel (€)", "Rotation Valeur (Proxy)",
                            "Forecast Ventes Période (IA)", "Ventes Moy Hebdo Prévues (IA)",
                            "WoS Projeté Début Période (IA)", "Stock Projeté Fin Période (IA)",
                            "Vts N-1 Tot (Mois Sel.)", "Qté Tot Prév (Mois Sel.)",
                            "Mnt Tot Prév (€) (Mois Sel.)"
                        ])
                        for month_name_cal in calendar.month_name[1:]:
                            known_non_week_cols_set.update([
                                f"Ventes N-1 {month_name_cal}", f"IA Ventes Brutes {month_name_cal}",
                                f"Qté Prév. {month_name_cal}", f"Montant Prév. {month_name_cal} (€)"
                            ])

                        for col_cand_sem in candidate_cols_sem:
                            col_cand_sem_str = str(col_cand_sem)
                            if col_cand_sem_str not in known_non_week_cols_set:
                                try:
                                    is_week_name_format = parse_week_column_to_date(col_cand_sem_str) is not None
                                    
                                    col_data_no_na = df_fs_loc_filter[col_cand_sem].dropna()
                                    is_numeric_col_content = False
                                    if not col_data_no_na.empty:
                                        if pd.api.types.is_numeric_dtype(col_data_no_na.dtype):
                                            is_numeric_col_content = True
                                        else:
                                            numeric_conversion_ratio = pd.to_numeric(col_data_no_na, errors='coerce').notna().mean()
                                            if numeric_conversion_ratio > 0.8:
                                                is_numeric_col_content = True
                                    
                                    if is_week_name_format or is_numeric_col_content:
                                        potential_sem_cols_read.append(col_cand_sem)
                                except Exception:
                                    pass 
                        st.session_state.all_available_semaine_columns = potential_sem_cols_read
                        if not potential_sem_cols_read: 
                            st.warning("⚠️ Aucune colonne de vente par semaine n'a été automatiquement identifiée.")
                    else:
                         st.warning("⚠️ Pas de colonnes trouvées après 'Désignation Article'.")
                except KeyError: 
                    st.warning("Colonne 'Désignation Article' non trouvée. La détection des colonnes de semaine est impossible.")
            else:
                if df_fs_loc_filter is None or df_fs_loc_filter.empty:
                    st.warning("Le 'Tableau final' n'est pas chargé, détection des semaines impossible.")
                elif "Désignation Article" not in df_fs_loc_filter.columns:
                    st.warning("Colonne 'Désignation Article' manquante, détection des semaines impossible.")
            st.info(f"Détection des colonnes semaine terminée: {len(st.session_state.all_available_semaine_columns)} colonnes identifiées.")

            progress_bar.progress(90, text="Préparation des listes de fournisseurs...")
            if not st.session_state.df_initial_filtered.empty and "Fournisseur" in st.session_state.df_initial_filtered.columns:
                st.session_state.unique_suppliers_list = sorted(
                    st.session_state.df_initial_filtered["Fournisseur"].astype(str).dropna().unique().tolist()
                )
            else: 
                st.session_state.unique_suppliers_list = []
                st.warning("Aucun fournisseur unique n'a pu être extrait.")
            st.info(f"Liste des fournisseurs uniques préparée: {len(st.session_state.unique_suppliers_list)} fournisseurs.")

            st.session_state.debug_last_file_processed_name = uploaded_file.name
            st.session_state.debug_df_full_was_set = True 
            progress_bar.progress(100, text="Chargement du fichier principal terminé !")
            st.success("✅ Fichier principal chargé et données de base préparées.")
            st.balloons()
            logging.info("Fichier principal traité avec succès. Appel de st.rerun().")
            st.rerun()

        except Exception as e_load_main_fatal_accueil:
            st.error(f"❌ Erreur majeure lors du chargement du fichier: {e_load_main_fatal_accueil}")
            logging.exception("Erreur majeure lors du chargement du fichier dans Accueil.py:")
            
            keys_to_clear_on_fatal_error = list(get_master_default_session_state().keys())
            preserved_keys_fatal = ['session_state_initialized'] 
            for pkf in preserved_keys_fatal:
                if pkf in keys_to_clear_on_fatal_error: keys_to_clear_on_fatal_error.remove(pkf)
            
            default_state_after_fatal_error = get_master_default_session_state()
            for key_to_reset_fatal in keys_to_clear_on_fatal_error:
                st.session_state[key_to_reset_fatal] = default_state_after_fatal_error[key_to_reset_fatal]
            
            st.session_state.df_full = None 
            st.session_state.debug_df_full_was_set = False
            st.session_state.debug_last_file_processed_name = None
            st.session_state.uploaded_file_name = None
            
            st.warning("L'état de session a été réinitialisé suite à une erreur. Rechargez un fichier valide.")
            logging.info("État de session réinitialisé suite à une erreur majeure. Appel de st.rerun().")
            st.rerun()

if not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour commencer.")
    if st.button("🔄 Réinitialiser l'Application (État Complet)", key="reset_app_button_accueil_no_file_v7"):
        current_file_uploader_state_key = "master_file_uploader_accueil"
        current_file_uploader_state = st.session_state.get(current_file_uploader_state_key, None)
        
        keys_before_reset = list(st.session_state.keys())
        for k_reset in keys_before_reset:
            if not k_reset.startswith("_stcore_") and k_reset != current_file_uploader_state_key :
                try:
                    del st.session_state[k_reset]
                except KeyError:
                    pass
        
        default_state_values_manual_reset = get_master_default_session_state()
        for key_reset_val, default_value_reset in default_state_values_manual_reset.items():
            if key_reset_val not in st.session_state:
                st.session_state[key_reset_val] = default_value_reset
        st.session_state.session_state_initialized = True
        
        if current_file_uploader_state is not None:
             st.session_state[current_file_uploader_state_key] = current_file_uploader_state
        
        st.success("L'état complet de l'application a été réinitialisé.")
        logging.info("Réinitialisation manuelle de l'application. Appel de st.rerun().")
        st.rerun()

elif st.session_state.get('df_full') is None or not st.session_state.get('debug_df_full_was_set', False):
    st.warning("Les données ne sont pas chargées. Veuillez (re)charger le fichier Excel principal.")

elif not st.session_state.get('df_initial_filtered', pd.DataFrame()).empty or \
     (st.session_state.get('df_suivi_commandes_all') is not None and not st.session_state.get('df_suivi_commandes_all').empty) or \
     (st.session_state.get('df_product_events') is not None and not st.session_state.get('df_product_events').empty):
    st.success(f"Fichier principal '{st.session_state.uploaded_file_name}' chargé. Données prêtes pour analyse.")
    st.markdown(f"""
    **Données disponibles :**
    *   Tableau Principal Filtré : `{len(st.session_state.df_initial_filtered)}` articles.
    *   Suivi des Commandes : `{len(st.session_state.df_suivi_commandes_all) if st.session_state.df_suivi_commandes_all is not None else 0}` lignes.
    *   Événements Spéciaux : `{len(st.session_state.df_product_events) if st.session_state.df_product_events is not None else 0}`.
    *   Fournisseurs uniques : `{len(st.session_state.unique_suppliers_list)}`
    *   Colonnes de Ventes Semaine détectées : `{len(st.session_state.all_available_semaine_columns)}`.
    """)
    if not st.session_state.all_available_semaine_columns:
        st.warning("⚠️ Attention : Aucune colonne de vente par semaine n'a été auto-détectée. Certaines analyses pourraient être limitées.")
    
    # --- DESCRIPTION DES APPLICATIONS DISPONIBLES (MODIFIÉE) ---
    st.markdown("---")
    st.subheader("Applications Disponibles")
    st.markdown("""
    Utilisez la **barre latérale de navigation** pour accéder aux outils d'analyse :

    - **1 - Commandes Stock** : Calcul des besoins de commande basés sur les ventes passées et les paramètres de stock.
    - **...** (vos autres applications)
    - **6 - RFA** : Suivi du chiffre d'affaires par fournisseur pour atteindre les paliers de remises de fin d'année.
    - **7 - Stocks (Scoring & Prévisions)** : Analyse avancée des stocks par fournisseur.
    - **8 - Statistiques Fournisseurs** : Vue d'ensemble et analyse détaillée de la performance des fournisseurs.
    - **9 - Réconciliateur de Documents** : **(Nouveau)** Comparez deux documents (Excel, PDF) pour identifier les écarts de prix et de quantité, même avec des références non-exactes.
    - **10 - Plans d'achat** : ** Suivi des plans d'achat et des stocks afférents.
    - **11 - Recherche BL** : ** Recherchez et exportez rapidement des Bons de Livraison (BL) à partir d'un fichier Excel dédié.
    - **12 - Commandes Fournisseur Urgentes** : ** Commandes Fournisseur à passer en urgence
    """)
    # --------------------------------------------------------

else:
    st.warning("Le fichier semble chargé, mais aucune donnée utile n'a pu en être extraite. Vérifiez son contenu et son format.")

# --- FIN DE Accueil.py ---
