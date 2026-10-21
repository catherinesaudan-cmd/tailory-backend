# À DÉPOSER — SERVEUR 2.73 ET PAGE 10.442 — préparé le 21.10.2026 à 16 h 08 (heure lue) ; le dépôt est à Catherine (DEC-3354, sa dérogation pour les anciennes batteries de la page)

**Ce fichier ne se dépose pas.**

## Ce qui change en ligne
| fichier | empreinte | en ligne aujourd'hui |
|---|---|---|
| `app.py` (= `app_v2_73.py` à l'octet) | 56281833e89af237 | 7171f5bbe9dcbef2 (la 2.53) → **change** |
| `EMPREINTES.txt` | — | → **change** (une ligne de plus : la police) |
| `police/Andika_400_mem.ttf` (dans un dossier `police`) | 00e4d4681327ebe4 | **absent → nouveau** (sans lui, ER s'abstient et le dit : les calculs restent en image) |
| `tailoryv10_442.html` (la page) | a11bbd6f167bd0a9 | la 10.414 → **change** |
| `grilles.py` · `grilles_v2_14.py` · `formes.py` · `formes_v2_10.py` · `requirements.txt` | voir `EMPREINTES.txt` | **déjà en ligne à l'octet** (comparés au dépôt GitHub le 21.10) : les déposer ne change rien |

## Les gestes, dans l'ordre
1. **Le serveur** : ouvrir `github.com/catherinesaudan-cmd/tailory-backend` → **« Add file » → « Upload files »**.
2. Glisser depuis ce dossier : **`app.py`**, **`EMPREINTES.txt`** et **le dossier `police`** (le dossier entier, pour qu'il arrive sous le nom `police/Andika_400_mem.ttf`). Ne pas glisser `LISEZMOI.md` ni la page.
3. **« Commit changes »**. Render reconstruit tout seul (quelques minutes ; le Dockerfile copie tout le dépôt, la police comprise).
4. Ouvrir **`https://tailory-backend-v3.onrender.com/health`** : il doit dire **`"version":"2.73"`**, **`"empreinte":"56281833e89a"`**, `"grilles":"2.14"`, `"formes":"2.10"`. Tant qu'il dit 2.53, le nouveau serveur n'est pas en ligne.
5. **La page** : mettre `tailoryv10_442.html` à la place de la page actuelle, à `classedobsdecath.ch/outils/tailory/`, sous le même nom que celle qui y est aujourd'hui.
6. Recharger la page : la pastille doit dire **Tailory v10.442**. Sur la première fiche générée, le pied **« VERSION LIVRÉE »** doit annoncer **10.442** et le serveur **2.73**.

**Entre le geste 3 et le geste 5**, la page 10.414 parle au serveur 2.73. ER efface alors les calculs des images sans que la 10.414 sache les réécrire, et un fichier Word est refusé. Il vaut mieux ne rien générer avant la fin du geste 6.

**Vérifié avant** : la 2.73 ne contient aucune écriture que le Python 3.11 du serveur refuserait (recherche faite le 21.10 ; la 2.53 en ligne, même recherche, 0) ; ses bibliothèques sont toutes dans `requirements.txt` en ligne. **Non vérifié** : la reconstruction par Render elle-même.
