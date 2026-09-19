"""CLI entry point: ``outlook-mcp`` and ``python -m outlook_mcp``."""

from __future__ import annotations

import logging
import os
import sys

from outlook_mcp.server import build_server


def main() -> None:
    # stdio servers must NEVER log to stdout — that's the JSON-RPC stream.
    logging.basicConfig(
        level=os.environ.get("OUTLOOK_MCP_LOG", "INFO"),
        stream=sys.stderr,
        format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    )
    log = logging.getLogger("outlook_mcp")

    mcp, bridge = build_server()
    log.info(
        "Starting outlook_mcp (stdio transport) - Outlook is attached "
        "lazily on the first tool call, not now"
    )

    # NOTE: bridge.start() is intentionally NOT called here. Attaching
    # runs Dispatch("Outlook.Application"), which makes DCOM launch
    # OUTLOOK.EXE — so starting eagerly opened Outlook every time this
    # server loaded, whether or not any Outlook tool was ever used.
    # OutlookBridge.call() attaches on demand instead.
    try:
        mcp.run()
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
