# -*- coding: utf-8 -*-
"""pptx_translate — format-preserving PowerPoint translation.

Public API::

    from pptx_translate.core import ProviderConfig, Translator, build_provider
"""

from . import core, fit  # noqa: F401

__version__ = "1.0.0"
__all__ = ["core", "fit", "__version__"]
