# À DÉPOSER — SERVEUR 2.78 ET PAGE 10.447 (CLÔTURE DU CHANTIER 1) — préparé le 22.10.2026 à 08 h 39 ; le dépôt est à Catherine (sa dérogation pour la garde : DEC-3385)

**Ce fichier ne se dépose pas.**

## Ce qui change en ligne
| fichier | empreinte | en ligne aujourd'hui |
|---|---|---|
| `app.py` (= `app_v2_78.py` à l'octet) | b99cb5628a25f07c | 56281833e89af237 (la 2.73) → **change** |
| `EMPREINTES.txt` | — | → **change** (la ligne de `app.py`) |
| `tailoryv10_447.html` (la page) | 30bbb951cc085b14 | la 10.442 → **change** |
| `grilles.py` · `grilles_v2_14.py` · `formes.py` · `formes_v2_10.py` · `requirements.txt` · `police/Andika_400_mem.ttf` | voir `EMPREINTES.txt` | **identiques au dépôt du 21.10** : les déposer à nouveau ne change rien |

## Les gestes, dans l'ordre
1. **Le serveur** : ouvrir `github.com/catherinesaudan-cmd/tailory-backend` → **« Add file » → « Upload files »**.
2. Glisser depuis ce dossier : **`app.py`** et **`EMPREINTES.txt`**. Ne pas glisser `LISEZMOI.md`.
3. **« Commit changes »**. Render reconstruit tout seul (quelques minutes).
4. Ouvrir **`https://tailory-backend-v3.onrender.com/health`** : il doit dire **`"version":"2.78"`** et **`"empreinte":"b99cb5628a25"`**. Tant qu'il dit 2.73, la reconstruction n'est pas finie.
5. **La page** : mettre `tailoryv10_447.html` à la place de la page actuelle, à `classedobsdecath.ch/outils/tailory/`, sous le même nom que celle qui y est aujourd'hui.
6. Recharger la page : la pastille doit dire **Tailory v10.447**. Sur la première fiche générée, le pied **« VERSION LIVRÉE »** doit annoncer **10.447** et le serveur **2.78**.

**Entre le geste 3 et le geste 5**, la page 10.442 parle au serveur 2.78 : sans danger (le champ neuf `lt` est ignoré par la 10.442), mais faire les deux à la suite.

## Ce qui a été vérifié avant
- les deux pièces : batteries des règles figées v4 **vertes** (226 serveur, 220 page) ; journaux à jour (`JOURNAL_BACKEND_v2_78.md`, `JOURNAL_v10_447.md`) ;
- la garde des accroches : 41 vertes, 3 échecs, 7 abstentions, 2 rouges, sans régression de l'outil : **dérogation déclarée par elle** (DEC-3385) ;
- `app.py` se lit en Python 3.9 (donc en 3.11, celui de Render) ; **aucune bibliothèque neuve** par rapport à la 2.73 ;
- **ce qui change pour l'élève** : le chantier 1 (§ 2 du nouveau plan) ; **le tirage de clôture suit le dépôt** (grille `CHANTIER_MESURES_CH1_05-10/TIRAGE_CLOTURE_CH1/GRILLE_TIRAGE_CLOTURE_CH1.md`).
