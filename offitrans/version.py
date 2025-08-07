"""
Version information for Offitrans
"""

__version__ = "1.0.0"
__version_info__ = (1, 0, 0)

# Build and release information
__author__ = "Offitrans Contributors"
__email__ = "offitrans@example.com"
__license__ = "MIT"
__url__ = "https://github.com/minglu6/Offitrans"
__description__ = "A powerful Office file translation tool library"


def get_version():
    """Get the version string"""
    return __version__


def get_version_info():
    """Get the version info tuple"""
    return __version_info__


def get_full_info():
    """Get full version and project information"""
    return {
        "version": __version__,
        "version_info": __version_info__,
        "author": __author__,
        "email": __email__,
        "license": __license__,
        "url": __url__,
        "description": __description__,
    }
