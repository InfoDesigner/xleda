import warnings

# Suppress all UserWarnings and DeprecationWarnings from third-party libraries
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)

from .main import wb  # noqa: E402


__all__ = ['wb']