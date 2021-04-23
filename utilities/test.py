import os

import document

if __name__=='__main__':
	path_temp, f = 'c:/users/jonathan/desktop/','test.docx'
	d = document.Document(path_temp, f)
	d._debug = True
	d.empty_document()
	d.save()
	os.system('start ' + path_temp + f)
