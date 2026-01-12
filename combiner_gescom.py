import json
import pandas as pd
from pathlib import Path
from datetime import datetime

def combiner_gescom_vers_excel():
    """
    Combine les fichiers gescom_ql.docoms.json et gescom_test.docoms.json
    et crée un fichier Excel des ventes avec les colonnes :
    ref, date, entite, codeProduit, RefProduit, libelleProduit, prixAchat, prixVente, qte
    """
    
    # Liste des fichiers à traiter
    fichiers_gescom = [
        "gescom_ql.docoms.json",
        "gescom_test.docoms.json"
    ]
    
    # Liste pour stocker toutes les lignes de ventes
    toutes_les_ventes = []
    
    # Parcourir chaque fichier
    for fichier in fichiers_gescom:
        fichier_path = Path(fichier)
        
        if not fichier_path.exists():
            print(f"Attention: Le fichier {fichier} n'existe pas")
            continue
        
        print(f"\nTraitement de {fichier}...")
        
        try:
            # Charger le fichier JSON
            with open(fichier, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            transactions_traitees = 0
            
            # Parcourir chaque transaction
            for transaction in data:
                # Extraire la date
                date_obj = transaction.get('date', {})
                try:
                    if isinstance(date_obj, dict) and '$date' in date_obj:
                        date_str = date_obj['$date']
                        # Gérer le cas où $date est un dict avec $numberLong
                        if isinstance(date_str, dict):
                            if '$numberLong' in date_str:
                                timestamp = int(date_str['$numberLong']) / 1000
                                date_parsed = datetime.fromtimestamp(timestamp)
                            else:
                                continue
                        else:
                            # Parser la date ISO 8601 string
                            date_parsed = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                        
                        date_formatee = date_parsed.strftime('%b %d, %Y %I:%M:%S %p')
                    else:
                        continue
                except Exception as e:
                    # Si erreur de parsing, passer à la transaction suivante
                    continue
                
                # Extraire l'ID de la transaction (ref)
                id_obj = transaction.get('_id', {})
                if isinstance(id_obj, dict) and '$oid' in id_obj:
                    ref = id_obj['$oid']
                else:
                    ref = ''
                
                entite = transaction.get('entite', '')
                
                # Parcourir chaque ligne de la transaction
                lignes = transaction.get('lignes', [])
                for ligne in lignes:
                    # Garder codeProduit tel quel (sans mapping)
                    code_produit = ligne.get('codeProduit', '')
                    
                    vente = {
                        'ref': ref,
                        'date': date_formatee,
                        'entite': entite,
                        'codeProduit': code_produit,
                        'RefProduit': code_produit,
                        'libelleProduit': ligne.get('libelleProduit', ''),
                        'prixAchat': ligne.get('prixAchat', 0.0),
                        'prixVente': ligne.get('prixVente', 0.0),
                        'qte': ligne.get('qte', 0.0)
                    }
                    toutes_les_ventes.append(vente)
                
                transactions_traitees += 1
            
            print(f"  -> {len(data)} transactions totales")
            print(f"  -> {transactions_traitees} transactions traitées")
        
        except Exception as e:
            print(f"Erreur lors du traitement de {fichier}: {str(e)}")
            continue
    
    # Créer un DataFrame pandas
    df = pd.DataFrame(toutes_les_ventes)
    
    # Afficher les statistiques
    print(f"\n" + "="*80)
    print(f"Total de lignes de ventes: {len(df)}")
    print(f"Nombre d'entités: {df['entite'].nunique()}")
    print(f"Entités: {df['entite'].unique().tolist()}")
    
    # Générer le nom du fichier Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fichier_excel = f"ventes_{timestamp}.xlsx"
    
    # Exporter vers Excel
    df.to_excel(fichier_excel, index=False, sheet_name='Ventes')
    
    print(f"\nFichier Excel créé avec succès: {fichier_excel}")
    print(f"Nombre total de lignes exportées: {len(df)}")
    print("="*80)
    
    return fichier_excel

if __name__ == "__main__":
    combiner_gescom_vers_excel()
