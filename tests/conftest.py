import os
import sys
from pathlib import Path

# Debe ir ANTES de importar bot.py para evitar que llame a input() / getpass()
os.environ.setdefault("PAMI_USER", "test_user")
os.environ.setdefault("PAMI_PASS", "test_pass")

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
