"""Mock Microsoft Graph -- Entra group membership.

NOT a BoondManager source: different authentication (OAuth2 client
credentials), different envelope (`value[]` + `@odata.nextLink`), different
contract. See docs/SPEC-DEVIATIONS.md #4.

Since 0.5.0, it also serves the Graph "files" surface (driveItem) -- a site,
its default library, empty at startup (see `drive.py`).
"""

from .app import app

__all__ = ["app"]
