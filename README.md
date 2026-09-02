# Pallet Flow

Application 100% front-end (HTML/CSS/JS, aucun backend) pour le suivi des palettes
par livraison et par transporteur, destinée à remplacer une base Access.

**Outil volontairement sans sauvegarde** : aucune donnée n'est écrite où que ce soit
(pas de `localStorage`, pas de fichier, pas de serveur). Tout reste en mémoire le
temps de la session et disparaît au rechargement ou à la fermeture de la page —
import, calcul, export, et c'est tout. Ce choix évite les soucis de fichiers qui
grossissent au fil des mois et de données obsolètes qui traînent dans le navigateur.

## Utilisation

Tout tient dans un **seul fichier autonome, `index.html`** (CSS, JS et la librairie
Excel sont intégrés dedans) : télécharge ce fichier et double-clique dessus, ou
ouvre-le depuis ton navigateur (`Fichier > Ouvrir...`). Aucun dossier annexe, aucune
connexion internet, aucun serveur requis — le fichier fonctionne partout où tu le
déplaces (clé USB, email, poste local...).

## Workflow

0. **Dashboard** : KPI (palettes sorties, retournées, solde global, transporteurs
   actifs, mouvements en attente de rapprochement) et solde par transporteur en un
   coup d'œil.
1. **Import ZM19** : export SAP contenant les sorties de palettes (colonnes utilisées :
   `Typ Pal`, `Document`, `Exp`, `Date`, `Uté stckage` en identifiant technique de ligne,
   `Rue` / `Vill` / `Nom 1` / `Lvré à` pour le détail destinataire).
2. **Import VTTK** : export SAP contenant, par n° de transport, l'itinéraire de
   transport (`Nº du transport`, `Itinéraire transport`). Seules les lignes n° de
   chargement (préfixe `970`) sont conservées ; les n° d'expédition (préfixe `960`)
   et les lignes sans itinéraire (chargement non finalisé) sont ignorées.
3. **Rapprochement** : relie chaque ligne ZM19 à son itinéraire VTTK via le n° de
   transport (`Exp` = `Nº du transport`), puis résout le transporteur via le
   référentiel (onglet *Référentiels*). Les nouveaux codes d'itinéraire découverts
   sont ajoutés automatiquement au référentiel, avec le nom déjà renseigné quand le
   code correspond à une correspondance connue (voir ci-dessous).
4. **Mouvements** : historique des sorties (avec ville/destinataire), filtrable par
   année / mois / transporteur / statut, exportable en Excel avec mise en forme.
5. **Retours transporteur** : saisie manuelle des retours de palettes (les extractions
   SAP ne couvrent que les sorties).
6. **Solde transporteurs** : sorties − retours par transporteur (tous types de palette
   confondus), exportable en Excel avec mise en forme (couleurs selon le signe du
   solde). Un solde positif = le transporteur doit encore rendre des palettes.

Les exports Excel (`.xls`) sont générés avec une mise en forme complète (en-têtes
colorés, lignes alternées, solde coloré par signe) — la librairie `xlsx` embarquée
gratuite ne sait pas écrire de styles dans un vrai `.xlsx`, donc l'app génère un
tableau HTML stylé reconnu nativement par Excel.

## Référentiels

- **Types de palettes** : quantité réelle de palettes représentée par chaque code
  `Typ Pal` (ex. `FE125` = 5 palettes Europe). Pré-rempli, éditable.
- **Transporteurs** : nom du transporteur associé à chaque code d'itinéraire
  transport. Pré-rempli à partir des correspondances connues :

  | Préfixe / code | Transporteur |
  |---|---|
  | `FRBERN*` | Bernard |
  | `FRDIS*` | Dispam |
  | `FRCEV*` | CEVA |
  | `FRCOQU*` | Coquelle |
  | `FROMNI*` | Omnitrans |
  | `FRPROM*` | Promodal |
  | `FRDTS*` | DTS |
  | `FRXPO*` | XPO |
  | `FRSEXP*` | Export |
  | `FRSLT*` | SLT |
  | `FRMAA1` | CEVA (Auchan Blanquefort) |
  | `FRMAA2` | Perrenot (Auchan Nîmes) |
  | `FRMAU1` | Omnitrans (Allotie St Vit) |
  | `FRMAU2` | Dispam (Allotie Vendargues) |
  | `FRMAU3` | Omnitrans (Auchan Cournon) |
  | `FRMAU4` | Dispam (Auchan Aix en Provence) |
  | `FRSTEF`, `FRM1/2/3/5/6/700*` | STEF (routes GMS) |

  Tout nouveau code découvert lors d'un import VTTK qui ne correspond à aucune de
  ces règles est ajouté au référentiel avec un nom vide (« à nommer ») — à compléter
  dans l'onglet *Référentiels*.

## Limites connues

- Aucune persistance : fermer l'onglet ou recharger la page vide tout — pense à
  exporter avant de partir. Ce n'est pas un bug, c'est le mode de fonctionnement
  voulu (outil portable, toujours vierge).
- Mono-utilisateur, pas de partage de données entre postes (chacun importe ses
  propres extractions SAP).
- Import Excel manuel (pas de connexion SAP directe).

Si le besoin d'historiser dans le temps ou de partager entre plusieurs personnes
apparaît, ça pourra évoluer vers une vraie base de données avec backend — mais ce
n'est plus l'outil "portable et vierge" décrit ici.
