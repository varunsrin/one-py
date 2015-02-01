"""
OneNote Object Model in Python
"""

from __future__ import absolute_import

from .onepy import (
	OneNote,
	Hierarchy,
	HierarchyNode,
	Notebook,
	SectionGroup,
	Section,
	Page,
	Meta,
	PageContent,
	Title,
	Outline,
	Position,
	Size,
	OE,
	InsertedFile,
	Ink,
	Image
	)

from .onmanager import ONProcess

__version__ = "0.2.1"