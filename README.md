# Veille Action Logement, In'li et CDC Habitat — registre des offres et alertes e-mail

Deux choses, en continu, sur [logement-actionlogement.fr](https://logement-actionlogement.fr/search),
[inli.fr](https://www.inli.fr/locations/offres/) et [cdc-habitat.fr](https://www.cdc-habitat.fr/recherche/location)
(sections « In'li » et « CDC Habitat » plus bas) :

1. **Un registre de toutes les offres publiées** (toute la France), avec toutes les informations de chaque
   annonce, sa date de mise en ligne et sa date de retrait, consultable dans un fichier **Excel**.
2. **Un e-mail à chaque nouvelle offre dans votre zone, à 1 100 € par mois maximum** : Paris (tous arrondissements),
   Vincennes, Saint-Mandé, Boulogne-Billancourt, Issy-les-Moulineaux, Neuilly-sur-Seine, Levallois-Perret,
   Charenton-le-Pont, Montreuil.
   L'e-mail reproduit l'annonce : photos, loyer et frais, caractéristiques, résidence, conditions d'accès,
   description et bouton vers la fiche.

## Comment ça marche

Le site s'appuie sur une API publique (`api.logement-actionlogement.fr`), sans compte ni clé. À chaque passage
(toutes les 20 minutes), le script :

1. récupère la liste complète des offres (`offers-overview`, page par page) ;
2. compare avec ce qu'il connaît déjà : **nouvelles offres**, **retraits**, **remises en ligne**,
   **modifications** (loyer, disponibilité, surface…) ;
3. lit la **fiche détaillée** de chaque nouvelle offre (`offer-showcase-details`) : étage, meublé, dépôt de
   garantie, DPE, public visé, règles de la résidence… ;
4. envoie un e-mail par nouvelle offre de la zone ;
5. met à jour l'Excel.

### Horodatages

| Information | Ce que c'est | Précision |
|---|---|---|
| Publiée le (site) | Date de publication affichée par le site | au jour près |
| Mise en ligne détectée le | Première vérification où l'offre est apparue | ~20 min |
| Retrait détecté le | Première vérification où l'offre a disparu | ~20 min |
| Vue en ligne pour la dernière fois | Dernière vérification où elle était présente | ~20 min |

Un retrait n'est **validé qu'après 2 vérifications complètes d'affilée** sans l'offre, et jamais lors d'une
vérification où l'API a mal répondu : pas de faux retraits si le site a un raté. Si l'offre revient, c'est
noté comme « Remise en ligne ».

## In'li

In'li n'a pas d'API : le script lit directement les pages du site, de la même façon qu'un visiteur.

- **Liste** : `inli.fr/locations/offres/?page=1…N`, 24 annonces par page, environ 700 annonces en tout.
  Chaque annonce est identifiée par sa référence (ex. `PRV-251606`).
- **Fiche détaillée** : loyer charges comprises et hors charges, provisions de charges, stationnement, dépôt de
  garantie, encadrement des loyers, pièces, chambres, surface, étage, équipements, DPE et GES, description, photos.
  In'li n'affiche pas de date de publication : la date de mise en ligne est celle de la première détection.
- **Zone d'alerte** : par code postal (In'li ne donne pas le code INSEE), mêmes villes qu'Action Logement.
- **Plafond** : 1 100 € par mois **charges comprises** (le prix « cc » affiché par In'li).
  Attention : In'li propose du logement intermédiaire, souvent plus cher. Le jour de la mise en place, aucune des 24
  offres In'li de la zone n'était sous 1 100 €. Pour un plafond propre à In'li, ajoutez la variable
  `INLI_ALERT_MAX_RENT` (par exemple `1300`).
- **Données** : `data/inli/` (même organisation que pour Action Logement) et Excel `offres-inli.xlsx`, publié au
  même endroit que l'autre.
- **Protection anti-robots** : si le site renvoie une page presque vide au script (moins de la moitié des annonces
  attendues), le passage In'li est abandonné sans rien modifier. Après plusieurs échecs, vous recevez un e-mail :
  ajoutez alors la variable `USE_BROWSER` = `true`, et le script lira les pages avec un vrai navigateur.

## CDC Habitat

CDC Habitat n'a pas d'API publique, mais sa page de résultats « Louer » contient, pour sa carte, la **liste complète**
des logements à louer de toute la France (environ 830) : numéro de lot, loyer charges comprises, catégorie, position GPS.

- **Liste** : une seule page lue à chaque passage, ce qui suffit pour repérer nouveautés et retraits.
- **Vignette** : pour chaque logement, une seule fois, la petite fiche de la carte donne la ville, le code postal et le
  lien de l'annonce. Pour une nouveauté, elle est lue tout de suite, avant de décider d'une alerte.
- **Fiche détaillée** : loyer charges comprises, charges générales et de chauffage, dépôt de garantie, facture
  d'énergie estimée, encadrement des loyers (Paris), catégorie (social, intermédiaire, loyer libre…), pièces, surface,
  étage, équipements, année de construction, chauffage, DPE et GES, description, photos, référence, téléphone de
  l'agence. Les fiches des logements déjà en ligne se remplissent au fil des passages (120 par passage, ceux proches
  de la zone d'abord), soit environ 4 heures pour tout le stock.
- **Zone d'alerte** : par code postal, mêmes villes que pour les autres sites. **Plafond** : 1 100 € charges comprises.
- **Places de parking** : enregistrées dans l'Excel mais jamais signalées par e-mail.
- **Données** : `data/cdc/` et Excel `offres-cdc-habitat.xlsx`, publié au même endroit que les autres.
- E-mails envoyés au nom de « Veille CDC Habitat ».

Un problème sur un site n'empêche jamais les autres de tourner.

## Second profil : GrandAngoulême

Les nouvelles offres des **38 communes de GrandAngoulême** (liste Insee de la communauté d'agglomération), à
**600 € par mois maximum**, sont envoyées à **une autre adresse**, sur les trois sites (Action Logement, In'li,
CDC Habitat). Les logements temporaires et les parkings en sont exclus, comme pour le profil principal.

- Adresse : secret `EMAIL_TO_ANGOULEME` (plusieurs adresses possibles, séparées par des virgules). Sans ce
  secret, le profil est désactivé.
- Plafond : variable `ANGOULEME_MAX_RENT` (défaut `600`).
- Le bouton `test-email` envoie aussi un mail de test à cette adresse.
- La liste des communes est dans `profils.py`.

## Format de stockage (et pourquoi)

| Fichier | Rôle |
|---|---|
| `data/offres/…/<identifiant>.json` | **Référence** : un fichier par offre, avec toutes les données brutes renvoyées par le site (liste + fiche détaillée) et les horodatages de suivi |
| `data/historique.csv` | Journal de tous les événements (mise en ligne, retrait, remise en ligne, modification) |
| `data/etat.json` | État général du suivi |
| `offres-action-logement.xlsx` | **Vue** régénérée à partir des fichiers ci-dessus |

L'Excel n'est pas la base de référence, volontairement. C'est un fichier binaire : écrit toutes les 20 minutes,
il risquerait de se corrompre sur une écriture interrompue et alourdirait vite le dépôt. Les fichiers texte, eux,
ne bougent que lorsqu'une offre change, gardent un historique lisible, et permettent de **régénérer l'Excel à tout
moment** (`python veille.py --export`) même si on le supprime ou si on change ses colonnes.

### Le fichier Excel

- **Offres** : une ligne par offre (en ligne et retirées), environ 80 colonnes : suivi, localisation, loyer et
  frais, logement, énergie, résidence, conditions d'accès, descriptions, lien cliquable. Les colonnes
  « Autre · … » reprennent tout champ du site sans colonne dédiée : rien n'est perdu. Filtres activés.
- **Zone d'alerte** : les mêmes colonnes, limitées à vos villes.
- **Historique** : tous les événements, du plus récent au plus ancien.
- **À propos** : compteurs et définition de chaque date.

Sur GitHub, il est publié dans **Releases → Données de la veille**, à une adresse qui ne change pas :
`https://github.com/<votre-compte>/<votre-dépôt>/releases/download/donnees/offres-action-logement.xlsx`

---

## Installation sur GitHub (gratuit, rien à laisser allumé)

### 1. Préparer l'envoi des e-mails

Il faut un compte qui **envoie** les alertes ; l'adresse qui les **reçoit** peut être n'importe laquelle.
Le plus simple est un compte **Gmail**, éventuellement créé exprès :

1. Activez la **validation en deux étapes** : [myaccount.google.com/security](https://myaccount.google.com/security).
2. Créez un **mot de passe d'application** : [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
   Les 16 lettres affichées, sans espaces, sont votre `SMTP_PASSWORD`.

### 2. Créer le dépôt

1. Créez un dépôt sur GitHub et envoyez-y tous ces fichiers (y compris le dossier caché `.github/`).
2. **Settings → Secrets and variables → Actions → Secrets** → *New repository secret*, un par un :

| Secret | Valeur |
|---|---|
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `465` |
| `SMTP_USER` | l'adresse Gmail qui envoie |
| `SMTP_PASSWORD` | le mot de passe d'application |
| `EMAIL_FROM` | l'adresse Gmail qui envoie |
| `EMAIL_TO` | l'adresse qui reçoit (plusieurs possibles, séparées par des virgules) |

### 3. Tester puis lancer

**Actions → Veille annonces Action Logement → Run workflow** :

1. `test-email` : vous recevez un exemple d'alerte construit avec une vraie annonce parisienne du moment.
   Regardez dans les **spams** la première fois.
2. `check` : premier vrai passage. Toutes les offres actuelles sont enregistrées **sans alerte** (plusieurs
   milliers : environ 2 400 rien qu'en dessous de 1 100 € de loyer). Les fiches détaillées sont lues 300 par passage, en commençant par votre zone : le stock complet
   est rempli en quelques heures.
3. Ensuite, tout est automatique toutes les 20 minutes.

**Public ou privé ?** Un dépôt **public** a des minutes illimitées, mais son contenu est visible de tous :
données d'offres déjà publiques, liste de vos villes, Excel. Vos secrets restent masqués. Un dépôt **privé**
dispose de 2 000 minutes par mois : réglez le `cron` de `.github/workflows/veille.yml` sur `*/30 * * * *`.

---

## En local (ordinateur, Raspberry Pi, serveur)

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env            # puis remplissez .env
python veille.py --test-email   # exemple d'alerte
python veille.py                # premier passage
```

Puis toutes les 15 minutes avec `crontab -e` :

```
*/15 * * * * cd /chemin/vers/alerte-action-logement && .venv/bin/python veille.py >> veille.log 2>&1
```

L'Excel est alors dans `export/offres-action-logement.xlsx`.

## Changer la zone d'alerte

La zone se règle par **codes INSEE** (pas les codes postaux), pour éviter les homonymes : Montreuil (93) n'est pas
Montreuil-sur-Mer, Neuilly-sur-Seine n'est pas Neuilly-sur-Marne. `751*` couvre les 20 arrondissements de Paris.
Pour modifier la liste, ajoutez la variable `ALERT_ZONES` (onglet **Variables**), par exemple :
`751*,94080,94067,92012,92040,92051,92044,94018,93048,94033` (ajout de Fontenay-sous-Bois).

## Changer le plafond de loyer des alertes

Par défaut, seules les offres dont le loyer affiché ne dépasse pas **1 100 € par mois** déclenchent un e-mail.
C'est le montant mensuel total affiché par le site, celui que filtre son option « loyer max ». Une offre sans
loyer indiqué déclenche quand même l'alerte, par prudence. Le registre et l'Excel, eux, gardent toutes les offres,
quel que soit leur loyer.

Pour changer le plafond, ajoutez la variable `ALERT_MAX_RENT` (onglet **Variables**), par exemple `1200`, ou `0`
pour ne plus avoir de plafond.

## Tous les réglages

| Variable | Défaut | Rôle |
|---|---|---|
| `ALERT_ZONES` | voir ci-dessus | Communes qui déclenchent un e-mail |
| `ALERT_MAX_RENT` | `1100` | Loyer maximum (€/mois) pour recevoir un e-mail ; `0` = pas de plafond |
| `INLI_ALERT_MAX_RENT` | = `ALERT_MAX_RENT` | Plafond propre à In'li (charges comprises) |
| `INLI_ALERT_POSTCODES` | mêmes villes | Codes postaux de la zone In'li (`75*` = tout Paris) |
| `INLI_MAX_DETAILS_PER_RUN` | `150` | Fiches In'li lues par passage |
| `INLI_ENABLED` | `true` | `false` pour désactiver In'li |
| `CDC_ALERT_MAX_RENT` | = `ALERT_MAX_RENT` | Plafond propre à CDC Habitat (charges comprises) |
| `CDC_ALERT_POSTCODES` | mêmes villes | Codes postaux de la zone CDC Habitat (`75*` = tout Paris) |
| `CDC_MAX_DETAILS_PER_RUN` | `120` | Logements CDC Habitat complétés (vignette + fiche) par passage |
| `CDC_ENABLED` | `true` | `false` pour désactiver CDC Habitat |
| `ALERT_EXCLUDE_PRODUCTS` | `TEMP` | Action Logement : pas d'e-mail pour les logements temporaires (foyers, résidences jeunes actifs, résidences sociales), qui restent dans l'Excel (colonne « Type de location ») ; `aucun` = alertes sur tout |
| `ALERT_ON_RELISTING` | `true` | E-mail aussi quand une offre retirée revient |
| `ALERT_ON_FIRST_RUN` | `false` | E-mails dès le premier passage pour les offres déjà en ligne |
| `SEARCH_BODY` | toute la France | Périmètre enregistré (même format que la recherche du site) |
| `MAX_DETAILS_PER_RUN` | `300` | Fiches détaillées lues par passage |
| `REMOVAL_CONFIRM_RUNS` | `2` | Vérifications complètes sans l'offre avant de valider un retrait |
| `MAX_ALERTS_PER_RUN` | `25` | E-mails max par passage (le reste part au suivant) |
| `MAX_PHOTOS` | `4` | Photos intégrées par e-mail |
| `REQUEST_DELAY` | `0.5` | Pause entre deux appels à l'API (secondes) |
| `USE_BROWSER` | `false` | Passer par un navigateur si l'API refuse les requêtes directes |
| `SMTP_SECURITY` | selon le port | `ssl` (465) ou `starttls` (587) |

## En cas de problème

- **Plus d'e-mails ni de mises à jour** : après 6 vérifications ratées d'affilée, le script vous envoie un e-mail
  d'avertissement avec l'erreur. GitHub vous prévient aussi par e-mail si une exécution échoue.
- **« L'API refuse les requêtes directes »** : le site a renforcé sa protection anti-robots. Ajoutez la variable
  `USE_BROWSER` = `true` : le script fera alors les mêmes appels depuis un vrai navigateur sans fenêtre.
- **Un e-mail n'est pas parti** (mot de passe changé…) : l'alerte reste en attente et repart au passage suivant.
- **Le site change son API** : renvoyez une capture HAR du site pour adapter le script.

## Bon usage

Un passage = 30 à 150 appels à l'API pour la liste paginée (selon la taille de page acceptée par le site), plus les fiches des nouvelles offres, toutes les
20 minutes, avec une pause entre chaque appel. C'est modéré ; évitez de descendre sous 15 minutes. Les photos ne
sont pas archivées : leurs liens expirent au bout de quelques heures. Elles sont donc intégrées directement dans
les e-mails.
