#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2006-2017 NV Access Limited, Manish Agrawal, Derek Riemer, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

from . import Window

class WordDocument(Window):
	"""
	An abstract class for WordDocuments which only contains script bindings and abstract code common to both Object Model and UI automation implementations for MS Word
	It uses the same original name for the WordDocument class ensuring that existing key bindings are not broken.
	"""

	__gestures = {
		"kb:control+[":"increaseDecreaseFontSize",
		"kb:control+]":"increaseDecreaseFontSize",
		"kb:control+shift+,":"increaseDecreaseFontSize",
		"kb:control+shift+.":"increaseDecreaseFontSize",
		"kb:control+b":"toggleBold",
		"kb:control+i":"toggleItalic",
		"kb:control+u":"toggleUnderline",
		"kb:control+=":"toggleSuperscriptSubscript",
		"kb:control+shift+=":"toggleSuperscriptSubscript",
		"kb:control+l":"toggleAlignment",
		"kb:control+e":"toggleAlignment",
		"kb:control+r":"toggleAlignment",
		"kb:control+j":"toggleAlignment",
		"kb:alt+shift+downArrow":"moveParagraphDown",
		"kb:alt+shift+upArrow":"moveParagraphUp",
		"kb:alt+shift+rightArrow":"increaseDecreaseOutlineLevel",
		"kb:alt+shift+leftArrow":"increaseDecreaseOutlineLevel",
		"kb:control+shift+n":"increaseDecreaseOutlineLevel",
		"kb:control+alt+1":"increaseDecreaseOutlineLevel",
		"kb:control+alt+2":"increaseDecreaseOutlineLevel",
		"kb:control+alt+3":"increaseDecreaseOutlineLevel",
		"kb:control+1":"changeLineSpacing",
		"kb:control+2":"changeLineSpacing",
		"kb:control+5":"changeLineSpacing",
		"kb:tab": "tab",
		"kb:shift+tab": "tab",
		"kb:NVDA+shift+c":"setColumnHeader",
		"kb:NVDA+shift+r":"setRowHeader",
		"kb:NVDA+shift+h":"reportCurrentHeaders",
		"kb:control+alt+upArrow": "previousRow",
		"kb:control+alt+downArrow": "nextRow",
		"kb:control+alt+leftArrow": "previousColumn",
		"kb:control+alt+rightArrow": "nextColumn",
		"kb:control+downArrow":"nextParagraph",
		"kb:control+upArrow":"previousParagraph",
		"kb:alt+home":"caret_moveByCell",
		"kb:alt+end":"caret_moveByCell",
		"kb:alt+pageUp":"caret_moveByCell",
		"kb:alt+pageDown":"caret_moveByCell",
		"kb:alt+shift+home":"caret_changeSelection",
		"kb:alt+shift+end":"caret_changeSelection",
		"kb:alt+shift+pageUp":"caret_changeSelection",
		"kb:alt+shift+pageDown":"caret_changeSelection",
		"kb:control+pageUp": "caret_moveByLine",
		"kb:control+pageDown": "caret_moveByLine",
		"kb:NVDA+alt+c":"reportCurrentComment",
	}
 
