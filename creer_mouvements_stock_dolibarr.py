import json
import pandas as pd
import requests
from pathlib import Path
from datetime import datetime
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Any
import logging

# Configuration Dolibarr
#DOLIBARR_URL = "https://admin.quinleysarlu.com"
DOLIBARR_URL="http://localhost/gescom"
DOLIBARR_TOKEN = "9f6f47942bc8f7ebf6909b27a7555c3d427fdb7f"

# Configuration de traitement
MAX_WORKERS = 2  # Réduit pour éviter HTTP 429 (Too Many Requests)
RETRY_ATTEMPTS = 3  # Nombre de tentatives en cas d'échec
REQUEST_DELAY = 0.2  # Délai entre requêtes augmenté (secondes)
RATE_LIMIT_DELAY = 5  # Délai en cas d'erreur 429 (secondes)

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('mouvements_stock.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def charger_mapping_entrepots():
    """Charge le mapping des entrepôts depuis entrepots.json"""
    try:
        with open('entrepots.json', 'r', encoding='utf-8') as f:
            entrepots = json.load(f)
        
        mapping = {}
        for entrepot in entrepots:
            ref = entrepot['ref']
            warehouse_id = int(entrepot['id'])
            mapping[ref] = warehouse_id
        
        return mapping
    except Exception as e:
        print(f"Erreur lors du chargement du fichier entrepots.json: {e}")
        return {}

# Mapping des entrepôts (entité -> warehouse_id Dolibarr)
WAREHOUSE_MAPPING = charger_mapping_entrepots()

# Configuration API
API_ENDPOINT = f"{DOLIBARR_URL}/api/index.php/stockmovements"
HEADERS = {
    "DOLAPIKEY": DOLIBARR_TOKEN,
    "Content-Type": "application/json"
}

def creer_mouvement_stock(product_id, warehouse_id, qty, movementcode="INV123", 
                          movementlabel="Inventory", price=None, datem=None, 
                          lot=None, dlc=None, dluo=None, retry_count=0):
    """
    Crée un mouvement de stock dans Dolibarr avec gestion des retries
    
    Args:
        product_id: ID du produit dans Dolibarr (required)
        warehouse_id: ID de l'entrepôt dans Dolibarr (required)
        qty: Quantité (positif = entrée, négatif = sortie) (required)
        movementcode: Code du mouvement (INV123, etc.)
        movementlabel: Label du mouvement
        price: Prix pour mettre à jour l'AWP
        datem: Date du mouvement
        lot: Numéro de lot
        dlc: Date limite de consommation (Eat-by date)
        dluo: Date limite d'utilisation optimale (Sell-by date)
        retry_count: Nombre de tentatives effectuées
    """
    
    payload = {
        "product_id": product_id,
        "warehouse_id": warehouse_id,
        "qty": qty,
        "movementcode": movementcode,
        "movementlabel": movementlabel
    }
    
    # Ajouter les champs optionnels s'ils sont fournis
    if price is not None:
        payload["price"] = price
    if datem:
        payload["datem"] = datem
    if lot:
        payload["lot"] = lot
    if dlc:
        payload["dlc"] = dlc
    if dluo:
        payload["dluo"] = dluo
    
    try:
        response = requests.post(API_ENDPOINT, headers=HEADERS, json=payload, timeout=30)
        
        if response.status_code == 200:
            return {
                "success": True,
                "id": response.json(),
                "message": "Mouvement créé avec succès"
            }
        else:
            # Gestion spéciale pour HTTP 429 (Too Many Requests)
            if response.status_code == 429:
                if retry_count < RETRY_ATTEMPTS:
                    wait_time = RATE_LIMIT_DELAY * (retry_count + 1)
                    logging.warning(f"⚠️  HTTP 429 (Rate Limit) - Attente de {wait_time}s avant retry...")
                    time.sleep(wait_time)
                    return creer_mouvement_stock(
                        product_id, warehouse_id, qty, movementcode, movementlabel, 
                        price, datem, lot, dlc, dluo, retry_count + 1
                    )
                else:
                    return {
                        "success": False,
                        "status_code": 429,
                        "message": "Rate limit dépassé après plusieurs tentatives"
                    }
            
            # Retry pour erreurs temporaires (500, 502, 503, 504)
            if response.status_code in [500, 502, 503, 504] and retry_count < RETRY_ATTEMPTS:
                time.sleep(1 * (retry_count + 1))  # Délai exponentiel
                return creer_mouvement_stock(
                    product_id, warehouse_id, qty, movementcode, movementlabel, 
                    price, datem, lot, dlc, dluo, retry_count + 1
                )
            
            return {
                "success": False,
                "status_code": response.status_code,
                "message": response.text
            }
    
    except requests.exceptions.Timeout:
        if retry_count < RETRY_ATTEMPTS:
            time.sleep(1 * (retry_count + 1))
            return creer_mouvement_stock(
                product_id, warehouse_id, qty, movementcode, movementlabel, 
                price, datem, lot, dlc, dluo, retry_count + 1
            )
        return {
            "success": False,
            "message": "Timeout après plusieurs tentatives"
        }
    
    except Exception as e:
        return {
            "success": False,
            "message": f"Erreur: {str(e)}"
        }

def convertir_date_vers_dolibarr(date_str):
    """
    Convertit une date du format 'Sep 3, 2025 12:00:00 AM' vers 'YYYY-MM-DD'
    """
    try:
        # Parser la date depuis le format d'origine
        date_obj = datetime.strptime(date_str, "%b %d, %Y %I:%M:%S %p")
        # Retourner au format YYYY-MM-DD
        return date_obj.strftime("%Y-%m-%d")
    except:
        # En cas d'erreur, retourner la date du jour
        return datetime.now().strftime("%Y-%m-%d")

def lire_excel_ventes(fichier_excel):
    """Lit le fichier Excel des ventes"""
    df = pd.read_excel(fichier_excel)
    return df

def lire_lignes_echouees(fichier_json):
    """Lit les lignes échouées depuis un fichier JSON"""
    with open(fichier_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
    df = pd.DataFrame(data)
    # Supprimer les colonnes techniques ajoutées
    df = df.drop(columns=['_erreur_precedente', '_ligne_originale'], errors='ignore')
    return df

def traiter_ligne_vente(index: int, row: Dict[str, Any], warehouse_mapping: Dict[str, int]) -> Dict[str, Any]:
    """
    Traite une ligne de vente individuelle (fonction pour parallélisation)
    
    Returns:
        Dict avec 'success', 'index', 'message', etc.
    """
    try:
        # Extraire les données
        ref = row['ref']
        date_vente = row['date']
        entite = row['entite']
        code_produit = row['codeProduit']
        libelle = row['libelleProduit']
        qte = row['qte']
        prix_vente = row['prixVente']
        
        # Mapper l'entrepôt
        warehouse_id = warehouse_mapping.get(entite)
        if not warehouse_id:
            return {
                "success": False,
                "index": index,
                "ref": ref,
                "erreur": f"Entité non mappée: {entite}"
            }
        
        # Utiliser codeProduit comme product_id Dolibarr
        try:
            product_id = int(code_produit)
        except:
            return {
                "success": False,
                "index": index,
                "ref": ref,
                "erreur": f"codeProduit invalide: {code_produit}"
            }
        
        # Créer le mouvement
        mouvement_code = f"FactMag-{ref}-{product_id}"
        mouvement_label = f"Facture magasin N° {ref}"
        date_mouvement = convertir_date_vers_dolibarr(date_vente)
        
        resultat = creer_mouvement_stock(
            product_id=product_id,
            warehouse_id=warehouse_id,
            qty=-abs(qte),
            movementcode=mouvement_code,
            movementlabel=mouvement_label,
            price=prix_vente,
            datem=date_mouvement
        )
        
        if resultat["success"]:
            return {
                "success": True,
                "index": index,
                "ref": ref,
                "product_id": product_id,
                "qte": -abs(qte),
                "mouvement_id": resultat['id']
            }
        else:
            return {
                "success": False,
                "index": index,
                "ref": ref,
                "erreur": resultat["message"]
            }
    
    except Exception as e:
        return {
            "success": False,
            "index": index,
            "ref": row.get('ref', 'N/A'),
            "erreur": f"Exception: {str(e)}"
        }

def creer_mouvements_depuis_excel(fichier_excel, mode="test", limite=10, est_retry=False):
    """
    Crée des mouvements de stock depuis un fichier Excel de ventes (version optimisée)
    
    Args:
        fichier_excel: Chemin vers le fichier Excel ou JSON d'échecs
        mode: "test" pour tester sur quelques lignes, "production" pour tout
        limite: Nombre de lignes à traiter en mode test
        est_retry: True si on retraite des lignes échouées
    """
    
    # Lire le fichier Excel ou JSON
    logging.info(f"Lecture du fichier {fichier_excel}...")
    
    if fichier_excel.endswith('.json'):
        df = lire_lignes_echouees(fichier_excel)
        est_retry = True
        logging.info(f"📋 Retry de lignes échouées précédemment")
    else:
        df = lire_excel_ventes(fichier_excel)
    
    logging.info(f"Total de {len(df)} lignes de ventes trouvées")
    
    # En mode test, limiter le nombre de lignes
    if mode == "test":
        df = df.head(limite)
        logging.info(f"Mode TEST: Traitement de {len(df)} lignes seulement")
    else:
        logging.info(f"Mode PRODUCTION: Traitement de toutes les {len(df)} lignes")
    
    # Statistiques
    resultats = {
        "succes": 0,
        "echecs": 0,
        "erreurs": [],
        "mouvements_crees": []
    }
    
    total_lignes = len(df)
    logging.info(f"\nCréation des mouvements de stock avec {MAX_WORKERS} threads parallèles...")
    print("-" * 80)
    
    # Traitement parallèle avec ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Soumettre toutes les tâches
        futures = {
            executor.submit(traiter_ligne_vente, index, row.to_dict(), WAREHOUSE_MAPPING): index 
            for index, row in df.iterrows()
        }
        
        # Traiter les résultats au fur et à mesure
        completed = 0
        for future in as_completed(futures):
            completed += 1
            resultat = future.result()
            
            if resultat["success"]:
                print(f"✓ [{completed}/{total_lignes}] Ligne {resultat['index']+1}: Mouvement {resultat['mouvement_id']} - Produit {resultat['product_id']}, Qté: {resultat['qte']}")
                logging.info(f"Succès ligne {resultat['index']+1}: Mouvement {resultat['mouvement_id']}")
                resultats["succes"] += 1
                resultats["mouvements_crees"].append(resultat['mouvement_id'])
            else:
                print(f"❌ [{completed}/{total_lignes}] Ligne {resultat['index']+1}: {resultat['erreur']}")
                logging.error(f"Échec ligne {resultat['index']+1}: {resultat['erreur']}")
                resultats["echecs"] += 1
                resultats["erreurs"].append({
                    "ligne": resultat['index'] + 1,
                    "ref": resultat.get('ref', 'N/A'),
                    "erreur": resultat['erreur']
                })
            
            # Petit délai pour éviter de surcharger l'API
            time.sleep(REQUEST_DELAY)
    
    # Afficher le résumé
    print("-" * 80)
    logging.info(f"\n📊 RÉSUMÉ")
    logging.info(f"Succès: {resultats['succes']}")
    logging.info(f"Échecs: {resultats['echecs']}")
    logging.info(f"Taux de réussite: {(resultats['succes']/total_lignes*100):.1f}%")
    
    if resultats['erreurs']:
        logging.warning(f"\n❌ ERREURS DÉTAILLÉES:")
        for err in resultats['erreurs'][:10]:  # Afficher les 10 premières
            logging.warning(f"  Ligne {err['ligne']}: {err.get('ref', 'N/A')} - {err['erreur']}")
    
    # Sauvegarder les IDs des mouvements créés
    if resultats['mouvements_crees']:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fichier_ids = f"mouvements_crees_{timestamp}.json"
        with open(fichier_ids, 'w') as f:
            json.dump(resultats['mouvements_crees'], f, indent=2)
        logging.info(f"\n✓ IDs des mouvements sauvegardés dans: {fichier_ids}")
    
    # Sauvegarder les lignes échouées pour retry ultérieur
    if resultats['erreurs']:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fichier_echecs = f"lignes_echouees_{timestamp}.json"
        
        # Récupérer les données complètes des lignes échouées
        lignes_echouees_data = []
        for err in resultats['erreurs']:
            ligne_num = err['ligne'] - 1  # Index pandas
            if ligne_num < len(df):
                row_data = df.iloc[ligne_num].to_dict()
                row_data['_erreur_precedente'] = err['erreur']
                row_data['_ligne_originale'] = err['ligne']
                lignes_echouees_data.append(row_data)
        
        with open(fichier_echecs, 'w', encoding='utf-8') as f:
            json.dump(lignes_echouees_data, f, indent=2, ensure_ascii=False)
        
        logging.warning(f"\n⚠️  {len(lignes_echouees_data)} lignes échouées sauvegardées dans: {fichier_echecs}")
        logging.warning(f"   Pour réessayer: python creer_mouvements_stock_dolibarr.py --retry {fichier_echecs}")
    
    return resultats

def tester_connexion_api():
    """Teste la connexion à l'API Dolibarr"""
    logging.info("Test de connexion à l'API Dolibarr...")
    logging.info(f"URL: {API_ENDPOINT}")
    
    try:
        response = requests.get(f"{DOLIBARR_URL}/api/index.php/status", headers=HEADERS, timeout=10)
        
        if response.status_code == 200:
            logging.info("✓ Connexion réussie à l'API Dolibarr")
            return True
        else:
            logging.error(f"❌ Échec de connexion: Status {response.status_code}")
            logging.error(f"Response: {response.text}")
            return False
    
    except Exception as e:
        logging.error(f"❌ Erreur de connexion: {str(e)}")
        return False

if __name__ == "__main__":
    import sys
    
    print("=" * 80)
    print("CRÉATION DE MOUVEMENTS DE STOCK DANS DOLIBARR (VERSION OPTIMISÉE)")
    print("=" * 80)
    
    # Vérifier si on est en mode retry
    fichier_a_traiter = None
    mode_retry = False
    
    if len(sys.argv) > 2 and sys.argv[1] == "--retry":
        fichier_a_traiter = sys.argv[2]
        mode_retry = True
        logging.info(f"\n🔄 MODE RETRY: Retraitement des lignes échouées")
        logging.info(f"📁 Fichier: {fichier_a_traiter}")
    
    # 1. Afficher la configuration
    logging.info(f"Configuration: {MAX_WORKERS} threads, {RETRY_ATTEMPTS} tentatives, délai {REQUEST_DELAY}s")
    
    # 2. Tester la connexion
    if not tester_connexion_api():
        logging.error("\n⚠️  Veuillez vérifier la configuration (URL, Token)")
        exit(1)
    
    # 3. Vérifier le mapping des entrepôts
    if not WAREHOUSE_MAPPING:
        logging.error("❌ Aucun entrepôt mappé. Vérifiez le fichier entrepots.json")
        exit(1)
    logging.info(f"Entrepôts mappés: {list(WAREHOUSE_MAPPING.keys())}")
    
    # 4. Trouver le fichier à traiter
    if not fichier_a_traiter:
        fichiers_excel = list(Path(".").glob("ventes_combinees_*.xlsx"))
        
        if not fichiers_excel:
            logging.error("\n❌ Aucun fichier Excel de ventes trouvé")
            logging.error("Veuillez d'abord exécuter combiner_ventes.py")
            exit(1)
        
        # Prendre le plus récent
        fichier_a_traiter = max(fichiers_excel, key=lambda p: p.stat().st_mtime)
    
    logging.info(f"\n📂 Fichier à traiter: {fichier_a_traiter}")
    
    # 5. Mode production activé
    if mode_retry:
        logging.info("\n🔄 MODE RETRY: Retraitement de toutes les lignes échouées")
    else:
        logging.info("\n🚀 MODE PRODUCTION: Traitement de toutes les lignes")
    logging.info("Démarrage du traitement...\n")
    
    start_time = time.time()
    
    # 6. Créer les mouvements
    resultats = creer_mouvements_depuis_excel(
        fichier_excel=str(fichier_a_traiter),
        mode="production",
        limite=10,
        est_retry=mode_retry
    )
    
    elapsed_time = time.time() - start_time
    logging.info(f"\n⏱️  Temps d'exécution: {elapsed_time:.2f} secondes")
    if resultats['succes'] > 0:
        logging.info(f"⚡ Vitesse: {resultats['succes']/elapsed_time:.1f} mouvements/seconde")
    logging.info("\n✓ Terminé")
