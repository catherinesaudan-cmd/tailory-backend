# RELAIS (diagnostic 29.07.2026, mis a jour 11.08.2026) — le serveur importe
# « grilles » ; le module livre s'appelle grilles_v2_12.py. Sans ce relais,
# l'import echoue en silence et le pipeline retombe en 2.9.3. A deployer
# A COTE de app.py sur Render, AVEC grilles_v2_12.py.
from grilles_v2_12 import *  # noqa
