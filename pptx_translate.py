# -*- coding: utf-8 -*-
"""Allow ``python pptx_translate.py <deck.pptx>`` as well as ``-m pptx_translate``."""
import sys

from pptx_translate.cli import main

if __name__ == "__main__":
    sys.exit(main())
