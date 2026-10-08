"""Mock Microsoft Graph — appartenance aux groupes Entra, et un dépôt SharePoint.

N'est PAS une source BoondManager : autre authentification (OAuth2 client
credentials), autre enveloppe (`value[]` + `@odata.nextLink`), autre contrat.
Cf. docs/SPEC-DEVIATIONS.md #4.

Depuis 0.5.0, la même application Entra lit aussi le dépôt Finance — un site
SharePoint, sa bibliothèque, un dossier de classeurs (cf. `drive.py`).
"""

from .app import app

__all__ = ["app"]
