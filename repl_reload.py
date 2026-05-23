"""repl_reload.py — Reload all project modules in bottom-up dependency order.

Call before each runpy invocation so edits to any module take effect:

    import repl_reload ; repl_reload.reload_all()
    sys.argv[1:] = ["scan"] ; import runpy ; temp = runpy._run_module_as_main("repl_main")
"""
import importlib

import word_tools

# Bottom-up: leaves first, entry point last.
RELOAD_ORDER = [word_tools]


def reload_all() -> None:
    importlib.invalidate_caches()
    for mod in RELOAD_ORDER:
        importlib.reload(mod)
