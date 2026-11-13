# -*- coding: utf-8 -*-
"""Compatibilidad temporal con el nombre V3; redirige al nuevo motor V5."""
from generar_ppt import main
if __name__=='__main__':
    import sys
    sys.exit(main())
