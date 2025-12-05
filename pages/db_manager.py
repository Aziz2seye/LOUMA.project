# Dans votre fichier Gestion_des_PVT.py, remplacez les fonctions de sauvegarde par :

from db_manager import ReportingDatabase

# Initialiser la base de données
if 'db_manager' not in st.session_state:
    st.session_state.db_manager = ReportingDatabase()
db = st.session_state.db_manager

# Charger les PVT depuis la base de données
pvt_df = db.get_all_pvt()

# Sauvegarder un nouveau PVT
def save_pvt(nom, contact):
    success, message = db.save_pvt(nom, contact)
    if success:
        st.success(message)
        st.session_state.pvt_data = db.get_all_pvt()  # Rafraîchir les données
        st.rerun()
    else:
        st.error(message)

# Mettre à jour un PVT
def update_pvt(old_nom, new_nom, new_contact):
    success, message = db.update_pvt(old_nom, new_nom, new_contact)
    if success:
        st.success(message)
        st.session_state.pvt_data = db.get_all_pvt()  # Rafraîchir les données
        st.rerun()
    else:
        st.error(message)

# Supprimer un PVT
def delete_pvt(pvt_nom):
    success, message = db.delete_pvt(pvt_nom)
    if success:
        st.success(message)
        st.session_state.pvt_data = db.get_all_pvt()  # Rafraîchir les données
        st.rerun()
    else:
        st.error(message)