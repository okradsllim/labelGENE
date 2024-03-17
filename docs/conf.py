# docs/conf.py

import os
import sys

# Project root directory as the parent of 'docs'
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Adding the project root directory to sys.path
sys.path.insert(0, project_root)

# Sphinx extension module names here, as strings.
extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.napoleon',  # Napoleon extension for Google-style docstrings
]

# Set autodoc options
autodoc_mock_imports = ['win32com']
autodoc_default_options = {
    'members': True,
    'undoc-members': True,
    'private-members': True,
    'special-members': '__init__',
    'imported-members': True,
}

# Master toctree document.
master_doc = 'index'

# General information
project = 'Your Project'
copyright = '2023, Will Nyarko'
author = 'Will Nyarko'

# Version and release
version = '0.5'
release = '1.0'

# directories to ignore when looking for source files.
exclude_patterns = ['_build']

# Pygments (syntax highlighting) style to use.
pygments_style = 'sphinx'

# Theme for HTML and HTML Help pages.
html_theme = 'sphinx_rtd_theme'