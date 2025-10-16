import sys
import types
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _ensure_stub_modules() -> None:
    """Provide lightweight stubs for optional dependencies during tests."""
    if "win32com" not in sys.modules:
        win32com_module = types.ModuleType("win32com")
        win32com_client = types.ModuleType("win32com.client")
        win32com_client.Dispatch = lambda *args, **kwargs: None
        win32com_module.client = win32com_client
        sys.modules["win32com"] = win32com_module
        sys.modules["win32com.client"] = win32com_client

    if "mcp" not in sys.modules:
        mcp_module = types.ModuleType("mcp")
        mcp_server_module = types.ModuleType("mcp.server")
        fastmcp_module = types.ModuleType("mcp.server.fastmcp")

        class DummyFastMCP:
            def __init__(self, *_, **__):
                self._tool_manager = types.SimpleNamespace(list_tools=lambda: [])
                self.settings = types.SimpleNamespace()

            def run(self, *args, **kwargs):  # pragma: no cover - unused in tests
                return None

            def tool(self, *args, **kwargs):
                def decorator(func):
                    return func

                return decorator

        fastmcp_module.FastMCP = DummyFastMCP
        fastmcp_module.Context = object

        exceptions_module = types.ModuleType("mcp.server.fastmcp.exceptions")

        class DummyToolError(Exception):
            pass

        exceptions_module.ToolError = DummyToolError

        mcp_module.server = types.SimpleNamespace(fastmcp=fastmcp_module)
        mcp_server_module.fastmcp = fastmcp_module

        sys.modules["mcp"] = mcp_module
        sys.modules["mcp.server"] = mcp_server_module
        sys.modules["mcp.server.fastmcp"] = fastmcp_module
        sys.modules["mcp.server.fastmcp.exceptions"] = exceptions_module


_ensure_stub_modules()
