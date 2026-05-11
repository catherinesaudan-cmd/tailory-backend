# Tailory Convert — Backend PDF→DOCX

## Déploiement sur Render.com (gratuit)

1. Crée un compte sur https://render.com
2. "New" → "Web Service"
3. Connecte ton GitHub et uploade ce dossier
4. Render détecte automatiquement le `render.yaml`
5. Clique "Deploy"
6. Note l'URL fournie (ex: `https://tailory-convert.onrender.com`)
7. Dans `tailory.html`, remplace `BACKEND_URL` par cette URL

## Endpoint

`POST /convert`
- Body : `multipart/form-data` avec un champ `file` (PDF)
- Retourne : fichier `.docx`

`GET /health`  
- Retourne : `{"status": "ok"}`
