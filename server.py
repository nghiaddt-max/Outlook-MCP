"""Loi vao chay thang bang python cua .venv (dung trong .mcp.json).
Ma nguon that nam trong goi outlook_mcp/."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from outlook_mcp.server import main  # noqa: E402

if __name__ == "__main__":
    main()
