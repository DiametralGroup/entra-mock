"""Container entrypoint: `python -m entra_mock`."""

from __future__ import annotations

import os


def main() -> None:
    import uvicorn

    uvicorn.run(
        "entra_mock:app",
        host=os.environ.get("ENTRA_MOCK_HOST", "0.0.0.0"),
        port=int(os.environ.get("ENTRA_MOCK_PORT", "8000")),
        log_config=None,
    )


if __name__ == "__main__":
    main()
