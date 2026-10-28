import sys
from core import get_db_path, init_db

if __name__ == "__main__":
    try:
        path = get_db_path()
        init_db(path)
        print(f"[OK] Base de datos verificada e inicializada en: {path}")
        sys.exit(0)
    except Exception as e:
        print(f"[ERROR] Fallo en inicializacion: {e}", file=sys.stderr)
        sys.exit(1)
