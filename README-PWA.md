# PWA COMPTA MOBILE — V1

Squelette PWA offline-first pour la saisie des dépenses en caisse vers le
classeur Excel OneDrive. Livré le 2026-09-15 (reprise du chantier après pause
du 2026-09-05). Depuis le 14/11 2026 (après-midi) : cible = le VRAI classeur
du compte outlook.com partagé avec fabrice.atmo@live.fr par lien « Peut
modifier » (mode workbookMode="share" : lien résolu en item via /shares,
écriture testée OK) ; la copie API du drive SPO du compte reste le repli
(workbookMode="path").

## SCHÉMA D'ARCHITECTURE

```
[caisse] -> [formulaire PWA (onglet Saisie)]
                 |
                 v
          [file IndexedDB (js/db.js, uuid anti-doublon)]
                 |
                 v
          [moteur sync FIFO (js/sync.js, dry-run par défaut)]
                 |
        +--------+--------------------------+
        |                                   |
  voie (a) OAuth                     voie (b) lien de partage
  (js/graph.js)                      (js/link.js)
  SPA + PKCE                         shareId = u!base64url(lien)
  PATCH range A:E                    GET content -> chirurgie zip
  + calculate Full                   (js/xlsx-surgery.js) -> PUT + If-Match
        |                                   |
        +--------+--------------------------+
                 |
                 v
  [classeur OneDrive « depenses fixe sept 2026-2027.xlsx » (Compta familiale,
   compte outlook.com, partagé « Peut modifier » ; repli : copie API drive SPO)]
     feuille journal RELSEPT26 — lignes 7-88, colonnes A-E uniquement
     -> totaux E89/D90 (plages SUM) alimentés automatiquement (O1)
     -> zéro insertion structurelle, zéro touch aux formules/essai/Feuil5
```

## FICHIERS

| Fichier | Rôle |
|---------|------|
| index.html | Shell 4 onglets : Saisie / File / Réglages / Diagnostic |
| manifest.webmanifest | Installabilité PWA |
| sw.js | Service worker : cache shell, jamais les API |
| icon.svg | Icône vectorielle |
| css/styles.css | UI mobile une main |
| js/config.js | Config persistée localStorage, dryRun=true PAR DÉFAUT |
| js/db.js | IndexedDB : file + meta |
| js/xlsx-surgery.js | Chirurgie zip/XML zéro perte (sans dépendance) |
| js/graph.js | Adaptateur OAuth Microsoft Graph |
| js/link.js | Adaptateur lien de partage OneDrive |
| js/sync.js | Moteur sync FIFO + backoff (30 s/2 min/10 min, cap 5) |
| js/app.js | Logique UI + diagnostic |

## DÉMARRAGE LOCAL (immédiat, sans auth)

    cd productions/standalone/compta_mobile/pwa
    python -m http.server 8000
    # puis ouvrir http://localhost:8000

La saisie, la file et le mode simulation fonctionnent sans aucune
configuration : dryRun=true par défaut, aucune écriture OneDrive n'est émise.

## PROCÉDURE VOIE (b) — LIEN DE PARTAGE (critère de reprise)

1. OneDrive : clic droit sur « depenses fixe sept 2026-2027 API.xlsx » ->
   Partager -> « Toute personne disposant du lien peut MODIFIER » -> copier.
2. PWA -> Réglages : coller le lien, mode « Lien de partage », Enregistrer.
3. PWA -> Diagnostic -> « 1. Tester la lecture (GET) » : doit lister les
   étapes OK (métadonnées, contenu, signature zip, journal).
4. PWA -> Diagnostic -> « 2. Tester l'écriture (PUT, octets identiques) » :
   ré-upload à l'identique — valide le PUT + If-Match (non garanti
   officiellement par Microsoft sur OneDrive perso).
5. Si OK : Réglages -> décocher « Mode simulation » -> saisir une vraie
   dépense de test -> vérifier dans Excel (ligne 7, colonnes A-E, totaux).

Sécurité : un lien « Peut modifier » donne un droit d'écriture COMPLET sur
le classeur à quiconque détient le lien. Ne pas diffuser ; révocable à tout
moment dans OneDrive (Paramètres des liens).

## PROCÉDURE VOIE (a) — OAUTH GRAPH (critère de reprise)

Prérequis (une fois, ~10-15 min, dans le navigateur de Fabrice) :
1. https://entra.microsoft.com -> Identités -> Applications -> Inscriptions
   d'applications -> Nouvelle inscription.
2. Type « Comptes personnels Microsoft + organisationnels » (endpoint common).
3. Plateforme « Application à page unique » -> Redirect URI : l'URL finale
   (local : http://localhost:8000/ ; prod : https://<compte>.github.io/...).
4. Copier l'ID d'application (client) -> PWA -> Réglages -> Client ID.
5. PWA -> Diagnostic -> « Tester la connexion Microsoft Graph » : consent,
   puis lecture du classeur + journal. Si OK, critère (a) validé.

## DÉPLOIEMENT GITHUB PAGES (recommandé, URL stable HTTPS)

1. Créer un dépôt (public ou private + Pages payant) -> pousser le dossier
   pwa/ à la racine (ou dans /docs).
2. Settings -> Pages -> Source : branche main / racine.
3. URL : https://<compte>.github.io/<repo>/ — ajouter cette URL exacte
   comme Redirect URI dans l'inscription d'app (étape 3 ci-dessus).
4. Mettre à jour sw.js (bump CACHE) à chaque déploiement.

## GARANTIES TECHNIQUES (invariants verrouillés dans le code)

- Lignes 7-88 uniquement (journalLastRow=88, JAMAIS sous les totaux E89/D90).
- Colonnes A-E uniquement, jamais F (case de Fabrice, décision O3).
- Aucune insertion structurelle : les lignes manquantes (41-88 absentes du
  XML d'origine) sont créées dans la plage existante, jamais en dessous.
- Chirurgie zéro perte : toutes les parts du zip recopiées octet pour octet
  sauf la part journal (méthode héritée de repare_f18_copie_api.py,
  vérifiée CRC côté Python) ; styles s385/386/387 (ligne 7) puis
  s388/389/390 ; dates en série Excel ; libellé inlineStr (sharedStrings
  intact) ; liaison externe 2023-2024 jamais touchée.
- Anti-doublon : uuid par saisie + allocation de ligne par relecture du
  classeur avant chaque push (source de vérité = le journal réel).
- OAuth : SPA + PKCE sans secret client ; tokens persistés en localStorage
  avec renouvellement silencieux par refresh_token (ouverture déjà connectée) ;
- Cible partagée (mode share) : le lien « Peut modifier » est résolu en item
  (drives/{driveId}/items/{id}) via /shares et l'API workbook vise cet item ;
  test no-op du 14/11 2026 sur le live (session + PATCH valeurs identiques) OK ;
- dryRun=true par défaut : badge SIMULATION visible en permanence.

## LIMITES CONNUES (V1)

- api.onedrive.com et graph.microsoft.com appliquent CORS : les tests
  doivent être lancés depuis http://localhost ou une origine HTTPS propre,
  PAS depuis file:// (ouvrir le fichier directement).
- PUT via lien de partage : fonctionnel en pratique sur OneDrive perso,
  non garanti officiellement — d'où le test dédié avant tout usage réel.
- La voie (b) ne propage PAS les recalculs (pas d'API calculate) : les
  totaux E89/D90 se recalculent à l'ouverture dans Excel/OneDrive web.
- Mois suivants : changer journalSheet dans Réglages (RELSEPT26 ->
  feuillets futurs) ; la V2 pourra dériver le nom du mois automatiquement.
- Pas de gestion multi-membres (V2) ; pas d'OCR tickets (hors scope V1).

## PLAN DE TEST (avant mise en production)

1. Local dry-run : 3 saisies -> vérifier badges « simulé · ligne N ».
2. Diagnostic GET (voie b) : 5 étapes OK.
3. Diagnostic PUT : valide le critère (b).
4. dry-run OFF : 1 dépense réelle de 0,01 € « TEST PWA » -> vérifier dans
  Excel : ligne correcte, colonnes A-E, F intacte, totaux reculés,
  formules intactes (E7, chaîne mensuelle) -> supprimer la ligne de test
  dans Excel manuellement si souhaité.
5. Hors-ligne : couper le réseau, saisir, rallumer -> sync auto (event
   online) -> vérifier push.
6. Redémarrer l'app (sw actif) : la file doit être conservée (IndexedDB).