# -*- coding: utf-8 -*-
"""Entrada principal. Ejecutar: python generar_ppt.py --excel ... --mdd ... --ddf ..."""
import sys
from oneshot.cli import main
if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        import logging
        logging.getLogger('oneshot').exception('[ERROR CRÍTICO] %s',exc)
        sys.exit(1)
