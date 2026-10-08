"""Mock Microsoft Graph — appartenance aux groupes Entra.

N'est PAS une source BoondManager : autre authentification (OAuth2 client
credentials), autre enveloppe (`value[]` + `@odata.nextLink`), autre contrat.
Cf. docs/SPEC-DEVIATIONS.md #4.

Depuis 0.5.0, il sert aussi la surface « fichiers » de Graph (driveItem) —
un site, sa bibliothèque par défaut, vide au démarrage (cf. `drive.py`).
"""

from .app import app

__all__ = ["app"]
