from flask import Flask, render_template, request, send_file # On importe Flask : (routes / pages / formulaires)
from docxtpl import DocxTemplate # permet de remplir le modèle word avec les variables du contexte
from datetime import datetime # pour la date
import pandas as pd  # poru lire le CSV
import re, json  #re pour simplifier la recherche dans le CSV et json  pour gérer le format json (utile pour les tableaux avec des valeurs différentes selon la zone)
from io import BytesIO # le tampon mémoire qui sert à générer le .docx sans avoir à créer un fichier dans le dur

app = Flask(__name__)#création de l'app Flask

app.config.update(TEMPLATE='TemplateCCTP.docx', CSV_SEP=';',MAX_ZONES=4 ) #on configure le nom du template word à remplir, ce qui sépare les infos du csv (en l'occurence un ;) et le nombre de zones max (car 4 zones possibles en VT)

# Il se trouve que les infos dans le csv AC et VT ont des noms différents donc on crée un dicitionnaire d'aliases pour les données qu'on va chercher l'ensemble des dénominations trouvables dans les 2 types de CSV.
FIELD_ALIASES = {
    'nom_projet': ['nom projet','nom du projet'],
    'type_installation': ['nature de la centrale','type installation','type de centrale'],
    'maitre_ouvrage': ['nom client','maître d’ouvrage','maitre d’ouvrage','client'],
    'ville': ['localisation','ville'],
    'adresse': ['adresse','adresse du site'],
    'puissance_kwc': ['puissance de la centrale','puissance zone totale','puissance install','puissande l’install','puissande l\'install','puissance installée'],
    'valorisation': ["valorisation de l'énergie produite",'mode de valorisation','valorisation'],
}

 #catalague des systèmes d'intégrations qu'on a référencés dans le template word. on les classe selon leur type ce qui facilite la sélection
INTEGRATIONS = {
    'membrane': ['EPC Solaire iNova PV Lite Tilt GC FE','SOPRASOLAR FIX EVO TILT','DOME SOLAR - ROOF SOLAR'],
    'leste': ['Novotegra toit plat III','ESDEC - FlatFix Wave'],
    'bac acier': ['DOME SOLAR - Helios B2','DOME SOLAR - Kogysun i+','DOME SOLAR - Fibro-Solar','DOME SOLAR - Ital-Solar','NOVOTEGRA - Minirails Paysage','JORISIDE - JoriSolar RS-EVO','JORISIDE - JoriSolar Opti Roof'],
    'tuiles': ['NOVOTEGRA - Top-fix petits éléments'],
    'ombriere': ['JORISIDE RS PARK 2','ADIWATT - Profil Evolution','SPORASOLAR - PARK','DOME SOLAR - HELIOS RC3']
}
# Pareil mais sans tri
PV_MODULES = {'Jinko JKM450N-54HL4R', 'DGMEC PV Series', 'Voltec A126'}
# Toujours pareil
INVERTERS = {'HUAWEI SUN2000-100 KTL-M2', 'HUAWEI SUN2000-215 KTL-H0', 'ENPHASE IQ8 AC-72-M-INT', 'SOLAREDGE 90K + OPTIMISEUR DE PUISSANCE SOLAREDGE S1000'}

# Types de zones ombrières qu’on sait reconnaître rapidement qui sert en particulier dans le 9.3.4	Fourniture et pose des coffrets AC avec les balises if has_ombrieres à différencier de des balises if Ombrieres qui sont le fruit d'un choix opéré à la fin du formulaire (lors do cochage de la section Présence d'ombrières et présence de hangars)
OMB_TYPES = ["OMB VL DOUBLE","OMB VL SIMPLE","OMB VL PORTIQUE","OMB PL","OMB BOIS VL SIMPLE","OMB BOIS VL DOUBLE"]
# Pareil
TOITURE_TYPES = ["TT LESTE SUD","TT LESTE E/W","TT SOUDE","TT BAC ACIER"]

#Cette fonction sert a lire un fichier CSV venant soit d’un upload soit d’un texte collé.Le but c’est de toujours recuperer les données proprement: en forcant un séparateur précis on évite que pandas transforme les cases vides en NaN , on garde des chaines vides "" au lieu d'écrire "Nan"
def parse_csv(src, from_text=False):
    opts = dict(sep=app.config['CSV_SEP'], header=None, dtype=str, keep_default_na=False, engine='python')# prepare toutes les options pandas pour lire un CSV avec ; en separateur et tout en texte
    return pd.read_csv(BytesIO(src.encode()), **opts) if from_text else pd.read_csv(src, **opts) # lit le csv depuis un texte (transformé en fichier en memoire) si from_text=True sinon lit direct le fichier uploadé

#Cette fonction sert a chercher la premiere valeur qui se trouve juste apres un mot cle donné (alias) dans le CSV. Elle verifie chaque ligne et chaque colonne, compare en ignorant les majuscules/minuscules et les espaces, et renvoie la valeur de la cellule suivante si ca correspond.
def find_first(df, aliases):
    if df.empty: return '' # si le tableau est vide on renvoie vide
    if isinstance(aliases, str): aliases = [aliases] # si un seul alias est donné on le met dans une liste
    aliases = [a.strip().casefold() for a in aliases if a and a.strip()] # on nettoie chaque alias et on passe en minuscule
    if not aliases: return '' # si la liste est vide on renvoie vide
    for row in df.itertuples(index=False): # on parcourt chaque ligne
        vals = [str(x).strip() for x in row] # on nettoie chaque valeur de la ligne
        for i in range(len(vals) - 1): # on regarde chaque cellule sauf la derniere (pas la denrière car on cherche le mot-clé qui se trouve être à gauche, donc si on cherche la dernnière on trouvera jamais de valeur)
            if any(alias in vals[i].casefold() for alias in aliases): # si un alias est trouvé dans la cellule
                return vals[i+1].strip() # on renvoie la cellule d'apres
    return '' # si rien trouvé on renvoie une chaine vide

 # Cette fonction sert a recuperer une valeur dans le CSV en passant par le systeme d'alias défini en haut du fichier. On donne un nom officiel (canonical_key) et ca va chercher toutes les variantes possibles dans FIELD_ALIASES.
def find_value(df, canonical_key):
    return find_first(df, FIELD_ALIASES.get(canonical_key, [canonical_key]))

# Recupere toutes les infos d'une zone precise dans le CSV et les renvoie sous forme de dictionnaire
def extract_zone(df, n):
    typ = find_first(df, [f"typologie zone {n}"])  # cherche le type de la zone n (1,2,3,4)
    if not typ: return None # si pasde type trouvé, la zone n'existe pas
    return {'name': f'Zone {n}','type': typ or '','puissance': find_first(df, [f"puissance zone {n}"]) or '','modules': find_first(df, [f"nombre panneaux zone {n}", f"nb panneaux zone {n}"]) or ''} #type, puissance, panneau trouvé

# Detecte automatiquement toutes les zones presentes dans le CSV et retourne leurs infos
def detect_zones(df):
    if df.empty or df.shape[1] < 2: return [] # si CSV vide ou moins de 2 colonnes, on renvoie une liste vide
    nums = sorted({int(m.group(1)) for cell in df.iloc[:, 1] if (m := re.search(r'zone\s*(\d+)', str(cell), re.I))}) # recupere tous les numeros de zones trouves dans la 2e colonne
    return [z for z in (extract_zone(df, n) for n in nums) if z][:app.config['MAX_ZONES']] # appelle extract_zone pour chaque numero et garde seulement les zones valides jusqu'a la limite MAX_ZONES de 4 zones

# Choisit rapidement la liste d'integrations possible en fonction du type de toiture ou structure
def pick_integration(typ):
    t = (typ or '').lower()
    if 'omb' in t:        return INTEGRATIONS['ombriere']   # ombrière
    if 'tui' in t:        return INTEGRATIONS['tuiles']      # tuiles
    if 'bac acier' in t:  return INTEGRATIONS['bac acier']   # bac acier
    if 'leste' in t:      return INTEGRATIONS['leste']       # lesté
    if 'membrane' in t or 'soude' in t:   return INTEGRATIONS['membrane']    # membrane
    return []
# Sélectionner 400/800V pour le paragraphe 9.3.2.	Fourniture et pose TBGT avec la tension demandée
def get_voltage(bt_mt):
    return "400V" if bt_mt == "BT" else ("800V" if bt_mt == "MT" else "Non défini")

# verifie si une case a cocher du formulaire a ete cochee (renvoie True si oui sinon False)
def to_bool(form, key):
    return form.get(key) == 'on' # dans les formulaires HTML une checkbox renvoie "on" si elle est cochee

# convertit une valeur texte en nombre flottant (float) en gerant les virgules et les valeurs vides
def _to_float(s):
    try:
        s = (s or '').replace(',', '.').strip()  # si s est None on le remplace par '', on change la virgule en point et on enleve les espaces
        return float(s) if s and s != '-' else 0.0  # si c'est pas vide et pas juste '-' on le transforme en float sinon on renvoie 0.0
    except ValueError:
        return 0.0  # si la conversion echoue on renvoie 0.0

# Pareil dns l'autre sens, on convertit un nombre décimal en entier
def _to_int(s):
    try:
        s = (s or '').strip()
        return int(s) if s and s != '-' else 0
    except ValueError:
        return 0

# calcule la somme totale des puissances et du nombre de modules sur toutes les zones
def compute_totals(zones):
    return (
        sum(_to_float(z.get('puissance')) for z in zones),  # additionne toutes les puissances converties en float
        sum(_to_int(z.get('modules')) for z in zones)  # additionne tous les modules convertis en int
    )

# lit un champ du formulaire qui contient un tableau en JSON et le converti en liste python (sert pour les tableaux du lot charpente et recap), sinon renvoie liste vide
def load_table_json(form, name):
    try: return json.loads(form.get(name, '[]'))  # recupere la valeur, si vide prend '[]', puis parse en JSON
    except json.JSONDecodeError: return []  # si le JSON est invalide on renvoie une liste vide

# nettoie les lignes des tabkeaux Présence ombrières et Présence  tableau en enlevant les espaces et en s'assurant que toutes les cles existent
def sanitize_rows(rows):
    for r in rows:  # pour chaque ligne du tableau
        for k in ('type','desc','modules','orient','incli','hbp'):  # pour chaque champ attendu
            r[k] = (r.get(k) or '').strip()  # recupere la valeur, met '' si None, enleve les espaces
    return rows  # renvoie le tableau nettoyé

@app.route('/') # definit la route racine qui affiche la page d'accueil pour envoyer un CSV,charge et renvoie la page HTML upload.html au navigateur
def upload():
    return render_template('upload.html')

#Après l’envoi du CSV ; elle lit le fichier si présent, prépare toutes les données et affiche le grand formulaire par zones
@app.route('/form', methods=['POST'])
def form():
    f = request.files.get('csv_file')  #Récupère le fichier envoyé depuis upload.html
    df = parse_csv(f) if f and f.filename.lower().endswith('.csv') else pd.DataFrame()  # Si on a bien un .csv, on le lit avec parse_csv pour obtenir un tableau exploitable ; sinon on part sur un DataFrame vide pour quand même afficher le formulaire
    zones = detect_zones(df)  # Détecte automatiquement “Zone 1”, “Zone 2”, etc. depuis le CSV
    csv_data = {k: find_value(df, k) for k in FIELD_ALIASES}  # Extrait les informations “projet” grace aux alias (ex. nom_projet, ville, adresse, puissance_kwc) qui correspondent aux balises word({{ nom_projet }}, {{ ville }}, etc.)

    ctx = {  # Construit le contexte envoyé au template formulaire.html ; il préremplit l’interface et transporte les données jusqu’à generate pour produire le Word
        'csv_text': ("\n".join(df.astype(str).agg(';'.join, axis=1))) if not df.empty else '',  # Version texte du CSV (séparateur ;)
        'zones': zones,  # Liste des zones détectées plus haut; utilisée pour afficher une colonne par zone dans la table de paramètres
        'zones_json': json.dumps(zones),  # Sérialisation JSON ; pour que generate relise exactement les mêmes zones sans devoir relire le CSV
        'panel_options': list(PV_MODULES),  #panneaux proposés dans les listes déroulantes ; le choix final remontera dans z['module'] et sera réutilisé dans le word
        'inverter_options': list(INVERTERS),  #Pareil
        'integration_options_per_zone': [pick_integration(z['type']) for z in zones],  # Pour chaque zoneon propose la liste de SI correspondant grâce à la liste établie dans INTEGRATIONS
        'latitude': request.form.get('latitude', ''),  # on demande de remplir la latitude
        'longitude': request.form.get('longitude', ''),  # Pareil
        'AC_VT': request.form.get('AC_VT', 'Autoconsommation'),  # Choix “Autoconsommation / Vente Totale”  et repris generate pour alimenter les balises word
        'bt_mt': request.form.get('bt_mt', 'BT'),  # Choix BT/MT; repris par generate pour labalise VOTRE_TENSION dans le document word et les balises if bt_mt == "MT" dans 9.3.2.	Fourniture et pose TBGT
        'ZONES': zones,  # Permet au word d'utiliser zones en majuscules
        'NB_ZONES': len(zones),  # Nombre total de zones utile dans le word pour afficher un bloc seulement s’il y a au moins une zone

        **csv_data # ajoute dans le contexte toutes les infos projet extraites du CSV via les alias (ex. nom_projet, ville, adresse, puissance_kwc) ; chaque clé correspond directement à une balise du modèle Word pour être remplacée automatiquement
    }

    return render_template('formulaire.html', **ctx)  # Affiche le formulaire complet déjà prérempli ; l’utilisateur valide ensuite vers generate, qui produira le document Word à partir de ce contexte

@app.route('/generate', methods=['POST'])  # Génaration du document Word à partir des données du formulaire
def generate():
    g = request.form.get  # On crée un alias pour simplifier l'accès aux données du formulaire

    # On récupère les informations saisies par l'utilisateur, si elles sont manquantes, on prend une valeur par défaut
    nom_projet = (g('nom_projet') or g('nom projet') or 'inconnu').strip()  # Nom du projet
    ville = (g('ville') or 'inconnue').strip()  # Ville du projet
    adresse = (g('adresse') or 'inconnue').strip()  # Adresse du projet
    zones = json.loads(g('zones_json', '[]'))  # On récupère les zones sous forme de JSON depuis le formulaire

    # On prépare un dictionnaire avec toutes les informations qu'on va insérer dans le modèle Word
    ctx = {
        'nom_projet': nom_projet,  #Le nom du projet
        'ville': ville,  # Pareil
        'adresse': adresse,  #Pareil
        'date': datetime.today().strftime('%d/%m/%Y'),  #Date du jour, formatée en français
        'latitude': g('latitude', '').strip(),  #Latitude
        'longitude': g('longitude', '').strip(),  #Longitude
        'AC_VT': g('AC_VT', 'Autoconsommation'),  #Choix entre Autoconsommation ou Vente Totale
        'VOTRE_TENSION': get_voltage(g('bt_mt', 'BT')),  # 400V ou 800V
        'liaison_terre_zones': list({g(f'zone-{i}-liaison_terre', '') for i in range(len(zones))}),  # Récupère la liaison à la terre pour chaque zone. On crée un ensemble pour éviter les doublons.
        'decouplage_zones': list({g(f'zone-{i}-decouplage', '') for i in range(len(zones))}),  #Pareil
        'has_paratonnerre': any(f'zone-{i}-paratonnerre' in request.form for i in range(len(zones))),  # Pareil à la différence qu'on ne vérifie pas que l'ensemble des zones ait le paramètre de renséigné mais qu'au mois une zone l'ait pour savori si on affichera la partie dans 8.3.7 Fourniture et pose des coffrets DC
        'coffretDC': any(f'zone-{i}-coffretDC' in request.form for i in range(len(zones))),
        'has_sdis_or_icpe': any( # Pareil qu'au dessus mais cette fois on cherche à voir si ICPE OU préconisations SDIS apparait au moins une fois (l'un ou l'autre) toujours dans l'ensemble des zones
            ('Préconisations SDIS' in request.form.getlist(f'zone-{i}-autres_specificites')) or
            ('ICPE' in request.form.getlist(f'zone-{i}-typologie_batiment'))
            for i in range(len(zones))
        ),
        'Ombrieres': to_bool(request.form, 'Ombrieres'),  # Vérifie si l'utilisateur a sélectionné "Ombrières" dans le formulaire. N'influe pas sur l'apparition du tableau dans le formulaire mais sur l'apparition du tableau dans le word au sein du lot charpente.
        'Hangars': to_bool(request.form, 'Hangars'),  # Pareil
        'travaux_rh': to_bool(request.form, 'travaux_rh'),  # Pareil mais pour la section Réseaux humides dans le lot VRD
        'ouvrages_retention': to_bool(request.form, 'ouvrages_retention'),  # Pareil mais pour savoir si le paragraphe à rédiger par les VRDistes doit apparaitre
        'KEEP_LOT_BORNES': 'keep_lot_bornes' in request.form,# Permet de vérifier si on a coché le lot des bornes de recharge pour décider si cette section sera incluse dans le document final.
        'KEEP_LOT_CHARPENTE': 'keep_lot_charpente' in request.form, #Pareil
        'KEEP_LOT_GROS_OEUVRE': 'keep_lot_gros_oeuvre' in request.form, #Pareil
        'KEEP_LOT_FONDATIONS_SPECIALES': 'keep_lot_fondations_speciales' in request.form, #Pareil
        'KEEP_LOT_HTA': 'keep_lot_hta' in request.form,  # Pareil
        'bridage_dynamique_enabled': to_bool(request.form, 'bridage_dyn'), #Pareil pour le bridage
        'bridage_dynamique_value': (g('bridage_dyn_value', '') or '').strip(),
    }

    for i, z in enumerate(
            zones):  # On boucle sur chaque zone, "i" représente l'index de la zone (de 0 à N) et "z" la zone elle-même
        for key in ('mode_valorisation', 'typologie_batiment', 'referentiel_technique',
                    'autres_specificites'):  # On boucle sur les clés spécifiques à chaque zone (par exemple, mode de valorisation, typologie du bâtiment, etc.)
            z[key] = request.form.getlist(
                f'zone-{i}-{key}')  # On récupère les choix faits dans le formulaire pour chaque zone et chaque clé, sous forme de liste brute. "getlist" permet de récupérer toutes les valeurs sélectionnées si plusieurs options sont possibles.
            z[f'{key}_display'] = ", ".join(z[key]) if z[
                key] else "Non défini"  # On crée une version lisible des choix sous forme de texte (séparé par des virgules), pour l'afficher proprement dans le document Word. Si la liste est vide, on affiche "Non défini" pour cette zone.

        z['integration'] = g(f'zone-{i}-integration', 'Non défini')  # Choix de l'intégration
        z['module'] = g(f'zone-{i}-module', 'Non défini')  # Choix du module photovoltaïque
        z['inverter'] = g(f'zone-{i}-inverter', 'Non défini')  # Choix de l'onduleur
        z['webdyn'] = (g(f'zone-{i}-webdyn', 'Aucun') or 'Aucun').strip()  # Type de supervision (Webdyn)
        z['paratonnerre'] = f'zone-{i}-paratonnerre' in request.form  # Présence de paratonnerre
        z['coffretDC'] = f'zone-{i}-coffretDC' in request.form
        z['bridage_enabled'] = (g(f'zone-{i}-bridage_enabled') is not None)  # Activation du bridage statique
        z['bridage_value'] = (g(f'zone-{i}-bridage_value', '') or '').strip() if z['bridage_enabled'] else '' # On récupère la valeur du bridage et on l'assigne directement si activé et non vide

    flat_types = [t for z in zones for t in (z.get('typologie_batiment') or []) if t]  # Rassemble toutes les typologies de bâtiment
    type_installation = ', '.join(sorted(set(flat_types)))  # On dédoublonne et on crée une liste des types d'installation
    has_ombrieres = any((z.get('type') or '') in OMB_TYPES for z in zones)  # Vérifie si des zones ont des ombrières
    has_toiture = any((z.get('type') or '') in TOITURE_TYPES for z in zones)  # Vérifie si des zones ont des toitures
    total_puiss, total_mod = compute_totals(zones)  # Calcule la puissance totale et le nombre de modules

    # On met à jour le contexte avec ces données globales
    ctx.update({
        'ZONES': zones,  # Liste des zones à inclure dans le document
        'NB_ZONES': len(zones),  # Nombre total de zones
        'type_installation': type_installation,  # Liste des types d'installation
        'has_ombrieres': has_ombrieres,  # Présence d'ombrières
        'has_toiture': has_toiture,  # Présence de toitures
        'TOTAL_PUISSANCE': total_puiss,  # Puissance totale
        'TOTAL_MODULES': total_mod,  # Nombre total de modules
        'SELECTED_INTEGRATION': [z.get('integration', '') for z in zones],  # Liste des intégrations sélectionnées par zone
        'SELECTED_MODULES': [z.get('module', '') for z in zones],  # Liste des modules sélectionnés par zone
        'SELECTED_INV': [z.get('inverter', '') for z in zones],  # Liste des onduleurs sélectionnés par zone
        'AUTOCONSOMMATION': g('AC_VT', 'Vente Totale'),  # Choix AC/VT
        'BT_MT': g('bt_mt', 'BT'),  # Choix BT/MT
        'puissance_kwc': total_puiss  # Puissance totale pour le CCTP
    })

    lines = []  # Liste des lignes pour le bridage
    for idx, z in enumerate(zones, 1):
        if z.get('bridage_enabled'):
            inv = (z.get('inverter') or '').strip() or z.get('name', f'Zone {idx}')  # Nom de l'onduleur ou zone
            v = (z.get('bridage_value') or '').replace(',', '.').strip()  # Valeur du bridage
            lines.append(
                f"- L’onduleur {inv} sera bridé à {v} kVA" if v else
                f"- L’onduleur {inv} fera l’objet d’un bridage statique (valeur à définir)"
            )

    # Mise à jour du contexte pour le bridage
    ctx.update({'has_bridage': bool(lines), 'bridage_lines': lines, 'bridage_paragraph': "\n".join(lines)})

    webdyn_simple = [f"Zone {i + 1}" for i, z in enumerate(zones) if(z.get('webdyn') or 'Aucun').strip() == 'Webdyn simple']  # Zones avec Webdyn simple
    webdyn_bridage = [f"Zone {i + 1}" for i, z in enumerate(zones) if (z.get('webdyn') or 'Aucun').strip() == 'Webdyn avec bridage dynamique']  # Zones avec Webdyn et bridage dynamique
    coffret_suivi = [f"Zone {i + 1}" for i, z in enumerate(zones) if (z.get('webdyn') or 'Aucun').strip() == 'Coffret de supervision ELUM']  # Zones avec coffret de supervision ELUM
    ctx.update({'webdyn_simple': webdyn_simple, 'webdyn_bridage': webdyn_bridage,'coffret_suivi': coffret_suivi})  # Mise à jour du contexte

    ctx.update({
        'OMB_TABLE': sanitize_rows(load_table_json(request.form, 'omb_table')),  # Données pour les ombrières
        'HANG_TABLE': sanitize_rows(load_table_json(request.form, 'hang_table'))  # Données pour les hangars
    })

    tpl = DocxTemplate(app.config['TEMPLATE'])  # On charge le modèle de document
    tpl.render(ctx)  # On remplace les balises par les données du contexte
    buf = BytesIO()  # On crée un tampon mémoire pour stocker le fichier généré
    tpl.save(buf)  # On enregistre le fichier dans le tampon
    buf.seek(0)  # On se positionne au début du tampon pour pouvoir l'envoyer

    # Envoi du fichier au navigateur pour téléchargement
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"CCTP_{ctx.get('nom_projet', 'projet')}.docx",  # Nom du fichier téléchargé
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"  # Type MIME pour un fichier Word
    )

# Point d’entrée : “python app.py” lance un petit serveur web en local (debug = True)
if __name__ == '__main__':
    app.run(debug=True)
