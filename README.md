# Pallet Flow

Application web (HTML/JS, sans backend) pour le suivi des palettes par livraison et
par transporteur, destinée à remplacer une base Access. Prototype mono-utilisateur :
les données sont stockées dans le navigateur (`localStorage`).

## Utilisation

Ouvrir `index.html` dans un navigateur, ou servir le dossier avec un serveur statique :

```bash
python3 -m http.server 8000
# puis ouvrir http://localhost:8000
```

Aucune connexion internet n'est nécessaire : la librairie de lecture Excel
([SheetJS](https://sheetjs.com)) est fournie localement dans `js/vendor/`.

## Workflow

1. **Import ZM19** : export SAP contenant les sorties de palettes (colonnes utilisées :
   `Typ Pal`, `Document`, `Exp`, `Date`, `Uté stckage` en identifiant technique de ligne).
2. **Import VTTK** : export SAP contenant, par n° de transport, l'itinéraire de
   transport (`Nº du transport`, `Itinéraire transport`). Les lignes sans itinéraire
   (chargement non finalisé) sont ignorées à l'import.
3. **Rapprochement** : relie chaque ligne ZM19 à son itinéraire VTTK via le n° de
   transport (`Exp` = `Nº du transport`), puis résout le transporteur via le
   référentiel (onglet *Référentiels*). Les nouveaux codes d'itinéraire découverts
   sont ajoutés automatiquement au référentiel avec un nom à renseigner.
4. **Mouvements** : historique des sorties, filtrable par année / mois / transporteur /
   statut, exportable en XLSX.
5. **Retours transporteur** : saisie manuelle des retours de palettes (les extractions
   SAP ne couvrent que les sorties).
6. **Solde transporteurs** : sorties − retours par transporteur (tous types de palette
   confondus), exportable en XLSX. Un solde positif = le transporteur doit encore
   rendre des palettes.

## Référentiels

- **Types de palettes** : quantité réelle de palettes représentée par chaque code
  `Typ Pal` (ex. `FE125` = 5 palettes Europe). Pré-rempli, éditable.
- **Transporteurs** : nom du transporteur associé à chaque code d'itinéraire
  transport. Se complète au fil des imports VTTK — renseigner le nom dans l'onglet
  *Référentiels* dès qu'un nouveau code apparaît.

## Limites connues (prototype)

- Mono-utilisateur, données locales au navigateur (pas de partage entre postes).
- Import Excel manuel (pas de connexion SAP directe).
- Pas de reprise de l'historique de l'ancienne base Access (démarrage à blanc).

Ces points pourront évoluer vers une vraie base de données partagée avec backend
si le besoin de plusieurs utilisateurs simultanés se confirme.
