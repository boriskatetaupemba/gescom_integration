import json
import pandas as pd
from pathlib import Path
from datetime import datetime

def combiner_ventes_json_vers_excel():
    """
    Combine tous les fichiers JSON de ventes et crée un fichier Excel
    avec les colonnes : ref, date, entite, codeProduit, RefProduit, 
    libelleProduit, prixAchat, prixVente, qte
    """
    
    # Liste des fichiers JSON à traiter
    fichiers_json = [
        "local_data DEVI.json",
        "local_data KENYA.json",
        "local_data LIKASI.json",
        "local_data MAG2.json"
    ]
    
    # Liste pour stocker toutes les lignes de ventes
    toutes_les_ventes = []
    
    # Parcourir chaque fichier JSON
    for fichier in fichiers_json:
        fichier_path = Path(fichier)
        
        if not fichier_path.exists():
            print(f"Attention: Le fichier {fichier} n'existe pas")
            continue
        
        print(f"Traitement de {fichier}...")
        
        try:
            # Charger le fichier JSON avec gestion de plusieurs encodages
            data = None
            for encoding in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                try:
                    with open(fichier, 'r', encoding=encoding) as f:
                        data = json.load(f)
                    print(f"  -> Encodage utilisé: {encoding}")
                    break
                except (UnicodeDecodeError, json.JSONDecodeError):
                    continue
            
            if data is None:
                print(f"  -> Impossible de décoder le fichier")
                continue
            
            # Parcourir chaque transaction
            for transaction in data:
                ref = transaction.get('ref', '')
                date = transaction.get('date', '')
                entite = transaction.get('entite', '')
                
                # Parcourir chaque ligne de la transaction
                lignes = transaction.get('lignes', [])
                for ligne in lignes:
                    vente = {
                        'ref': ref,
                        'date': date,
                        'entite': entite,
                        'codeProduit': ligne.get('codeProduit', ''),
                        'RefProduit': ligne.get('RefProduit', ''),
                        'libelleProduit': ligne.get('libelleProduit', ''),
                        'prixAchat': ligne.get('prixAchat', 0.0),
                        'prixVente': ligne.get('prixVente', 0.0),
                        'qte': ligne.get('qte', 0.0)
                    }
                    toutes_les_ventes.append(vente)
            
            print(f"  -> {len(data)} transactions traitées")
        
        except Exception as e:
            print(f"Erreur lors du traitement de {fichier}: {str(e)}")
            continue
    
    # Créer un DataFrame pandas
    df = pd.DataFrame(toutes_les_ventes)
    
    # Vérifier si des données ont été collectées
    if len(df) == 0:
        print("\nAucune donnée de vente n'a pu être extraite des fichiers JSON.")
        print("Veuillez vérifier l'intégrité des fichiers.")
        return None
    
    # Afficher les statistiques
    print(f"\nTotal de lignes de ventes: {len(df)}")
    print(f"Nombre d'entités: {df['entite'].nunique()}")
    print(f"Entités: {df['entite'].unique().tolist()}")
    
    # Générer le nom du fichier Excel avec timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fichier_excel = f"ventes_combinees_{timestamp}.xlsx"
    
    # Exporter vers Excel
    df.to_excel(fichier_excel, index=False, sheet_name='Ventes')
    
    print(f"\nFichier Excel créé avec succès: {fichier_excel}")
    print(f"Nombre total de lignes exportées: {len(df)}")
    
    return fichier_excel

if __name__ == "__main__":
    combiner_ventes_json_vers_excel()
