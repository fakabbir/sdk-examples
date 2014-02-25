#!/usr/bin/python

# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

LIBREOFFICE = "soffice"

import os
import sys
import time
import uno

from com.sun.star.lang import Locale

class SpellChecker:
	'''
		LibreOffice Spell Checking example.

		This example runs outside of LibreOffice. But we still need
		LibreOffice installed and running.

		See http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XSpellChecker.html#details
	'''

	def __init__(self):

		# Get the uno component context from the PyUNO runtime.
		# See http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1uno_1_1XComponentContext.html
		local_ctx = uno.getComponentContext()

		# Create the UnoUrlResolver.
		resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)

		ctx = self.get_context(resolver)

		if not ctx:
			print('Cannot load LibreOffice.')
			sys.exit(0)

		# Gets the service manager instance to be used.
		smgr = ctx.ServiceManager

		# Create an instance of Spell Checker.
		# See http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1SpellChecker.html
		self.spellchecker = smgr.createInstanceWithContext("com.sun.star.linguistic2.SpellChecker", local_ctx)

	def get_context(self, resolver):
		'''
			Connect to LibreOffice.
			Start LibreOffice if needed.
			We need a connection to use Lightproof SpellChecker.
		'''

		ctx = None
		load = False
		times = 0
		pipe_name = 'spellcnn'

		while not load and times < 5:
			try:
				#ctx = resolver.resolve("uno:socket,host=localhost,port=42424;urp;StarOffice.ComponentContext")
				ctx = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name)
				load = True
			except Exception as e:
				print(e)
				print ('Starting LibreOffice...')
				#os.system(LIBREOFFICE + ' --accept="socket,host=localhost,port=42424;urp;StarOffice.ServiceManager" --headless --nologo --nofirststartwizard &')
				os.system(LIBREOFFICE + ' --accept="pipe,name=%s;urp;StarOffice.ServiceManager" --headless --nologo --nofirststartwizard &' % pipe_name)
				times += 1
				time.sleep(5)

		return ctx

	def is_valid(self, lang, word):
		''' Checks if a word is spelled correctly in a given language. '''
		loc = Locale(lang[0:2], lang[3:5], '')
		return self.spellchecker.isValid(word, loc, ())

	def spell(self, lang, word):
		'''
			None if word is spelled correctly using lang.
			Otherwise, if available, proposals for spelling alternatives will be returned.
		'''
		loc = Locale(lang[0:2], lang[3:5], '')
		sug = self.spellchecker.spell(word, loc, ())

		if sug:
			return sug.getAlternatives()
		else:
			return None

if __name__ == '__main__':

	print('SpellChecker example\n')

	S = SpellChecker()

	words = [{'lang' : 'en-US', 'word' : 'worud'},
			 {'lang' : 'en-US', 'word' : 'astronoma'},
			 {'lang' : 'pt-BR', 'word' : 'programmador'},
			 {'lang' : 'pt-BR', 'word' : 'conhesimento'}]

	for w in words:
		k = w['lang']
		v = w['word']
		ok = S.is_valid(k, v)

		print('"%s" word "%s" is%sOK.' % (k, v, ' ' if ok else ' not '))

		sug = S.spell(k, v)

		# An empty list mean no suggestions for the word given.
		# So we're assuming that the word is well formed.
		if sug:
			print('Suggestions: %s\n' % (', '.join(sug)))
