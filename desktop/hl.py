import pymupdf
from pdf2excel.core.visibility import *
d=pymupdf.open(); p=d.new_page()
p.insert_text((60,80),"Resaltado importante",fontsize=12)
r=p.search_for("Resaltado importante")[0]
a=p.add_highlight_annot(r); a.update()
print('visible before bake', [w[4] for w in visible_words(p)])
d.bake()
p=d[0]
print([(x['type'],x['fill'],x.get('fill_opacity'),x.get('blendmode'),x['rect']) for x in p.get_drawings()])
print('visible after bake', [w[4] for w in visible_words(p)])
print(p.get_text('words')[:2])
