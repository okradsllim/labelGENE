# Configuration file for the Sphinx documentation builder.

extensions = ['sphinx.ext.autodoc']

# Suffix of source filenames.
source_suffix = '.rst'

# Master toctree document.
master_doc = 'index'

# General info
project = 'labelGENE'
copyright = '2023, Will Nyarko'
author = 'Will Nyarko'

# Version and release
version = '0.5'
release = '0.5'

exclude_patterns = ['_build']

# Pygments (syntax highlighting)
pygments_style = 'sphinx'

# HTML theme
html_theme = 'sphinx_rtd_theme'