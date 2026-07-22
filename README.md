# Générateur de plannings IMSR (ECF-CERCA)

Application Streamlit réutilisable pour produire, en quelques minutes, un
planning Excel dans le style de votre modèle
« Planning_contrat_pro_TP_ECSR », pour :

- **Formations longues** : TP ECSR, TFP FMESR, CQP RUESRC, CQP AGEC
  → une feuille Excel par année scolaire (calendrier Septembre → Août),
  fidèle au modèle fourni (codes couleur Centre / FOAD / Entreprise / Examens,
  jours fériés grisés, récapitulatif d'heures).
- **Formations courtes** : stage de réactualisation, habilitation post-permis,
  formation continue → un planning compact sur la durée choisie (dates libres).

Une feuille **Modules** liste, dans les deux cas, les CCP/BC/UC de l'action
et leur volume horaire, avec calcul automatique du total.

## Déploiement (comme vos applications AALC)

1. Créez un nouveau dépôt GitHub (ex. `imsr-planning-generator`) et déposez-y
   les trois fichiers : `app.py`, `planning_engine.py`, `requirements.txt`.
2. Sur [streamlit.io/cloud](https://streamlit.io/cloud), créez une nouvelle
   application pointant vers ce dépôt et le fichier `app.py`.
3. L'application est alors accessible en ligne, comme vos outils AALC.

## Utilisation locale (facultatif)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fonctionnement

1. **Informations générales** : famille de formation, type d'action, nom de
   l'apprenant, type de contrat, dates de début/fin.
   - Pour les formations longues, la date de fin détermine automatiquement le
     nombre de feuilles « année scolaire » générées.
   - Pour les formations courtes, dates de début/fin totalement libres.
2. **Modules** : tableau éditable listant les modules de l'action.
   - Le **code** proposé dépend automatiquement du type de formation choisi à
     l'étape 1 : `CCP1` / `CCP2` pour un TP ECSR, `BC1` / `BC2` / `BC3` pour un
     TFP FMESR, un CQP RUESRC ou un CQP AGEC (liste modifiable).
   - La **modalité** du module (Présentiel / Distanciel) se choisit dans une
     liste déroulante.
   - Le volume horaire est saisi en deux colonnes distinctes : **heures de
     formation** et **heures de stage**, avec un total calculé automatiquement.
3. **Directives récurrentes** : pour les règles qui se répètent chaque
   semaine (ex. « CCP2 en distanciel tous les jeudis »). Une directive
   = un jour de semaine + une modalité +, en option, un module lié (parmi
   les codes saisis à l'étape 2) + une plage de dates. C'est désormais le
   **seul mécanisme de remplissage** du calendrier (mais jamais prioritaire
   sur les jours fériés). Quand un module est renseigné, les heures
   correspondantes sont automatiquement cumulées et affichées dans l'onglet
   Modules du fichier généré, en regard du volume horaire prévu (avec un
   écart calculé).
4. **Remarques et structuration** : un espace de texte libre pour noter la
   structuration générale du planning (grandes périodes, alternance, points
   de vigilance...). Ce texte est repris tel quel dans la zone
   « COMMENTAIRES » du fichier Excel généré. Un panneau optionnel permet
   d'importer un planning Excel déjà réalisé pour le consulter à l'écran
   pendant la saisie (uniquement à titre de repère — il n'est pas utilisé
   dans le fichier généré).
5. **Génération** : un clic produit le fichier Excel (.xlsx) téléchargeable,
   avec le récapitulatif d'heures (Centre/FOAD/Examens/Entreprise + TOTAL)
   calculé par formules, comme dans le modèle d'origine.

## Points à garder en tête

- Le remplissage des jours du calendrier repose uniquement sur les
  **directives récurrentes** (étape 3, ex. « CCP2 en FOAD tous les
  jeudis »). Les jours non couverts par une directive restent vides et se
  complètent ensuite directement dans Excel si besoin.
- Les cases du calendrier restent en **Centre / FOAD / Entreprise / Examens**
  (comme votre modèle actuel) ; le lien avec les modules se fait uniquement
  via l'onglet Modules et son total d'heures, pas case par case.
- Les polices utilisées (Calibri) sont volontairement standard afin de
  s'afficher correctement sur tous les postes ; les couleurs reprennent
  exactement celles de votre modèle (vert Centre, orange FOAD, bleu
  Entreprise, gris jours fériés).
- Pour une formation longue démarrant par exemple en janvier, l'outil crée
  quand même une feuille « année scolaire » complète (septembre → août) :
  seuls les jours antérieurs à la date de début restent vides.
