import importlib.machinery
import importlib.util
import sys
from pathlib import Path


def load_client():
    name = 'folder_grant_standalone_client'
    if name in sys.modules:
        return sys.modules[name]
    path = Path(__file__).parents[1] / 'Folder Grant API Client.pyw'
    loader = importlib.machinery.SourceFileLoader(name, str(path))
    spec = importlib.util.spec_from_loader(name, loader)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        loader.exec_module(module)
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module
