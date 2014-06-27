"""
Tests & Stuff
"""

import unittest
from onepy import OneNote

on = OneNote()

#lists all sections & notebooks open in onenote


class TestOneNote(unittest.TestCase):

	def test_instance(self):
		for nbk in on.hierarchy:
			print (nbk)
			if nbk.name == "SoundFocus":
				for s in nbk:
			 		print ("  " + str(s))
			 		for page in s:
			 			print ("        " + str(page.name.encode('ascii', 'ignore')))


		#print(on.get_page_content("{37B075B6-358E-04DA-193E-73D0AD300DA3}{1}{B0}"))		

		self.assertEqual(True, True)

if __name__ == '__main__':
    unittest.main()