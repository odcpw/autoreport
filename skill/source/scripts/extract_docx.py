#!/usr/bin/env python3
"""Extract DOCX paragraph/table boundaries and links; no OCR or anonymisation."""
import argparse,json,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def extract(source):
 with zipfile.ZipFile(source) as z:
  if sum(x.file_size for x in z.infolist())>200*1024*1024:raise ValueError('DOCX exceeds extraction size allowance')
  rels={}
  if 'word/_rels/document.xml.rels' in z.namelist():
   for x in ET.fromstring(z.read('word/_rels/document.xml.rels')):rels[x.get('Id')]=x.get('Target')
  root=ET.fromstring(z.read('word/document.xml'));body=root.find(W+'body');blocks=[]
  def paragraph(el,where):
   t=''.join(n.text or '' if n.tag==W+'t' else '\t' if n.tag==W+'tab' else '\n' if n.tag in (W+'br',W+'cr') else '' for n in el.iter())
   links=[{'text':''.join(x.itertext()),'target':rels.get(x.get(R+'id')),'anchor':x.get(W+'anchor')} for x in el.iter(W+'hyperlink')]
   images=sum(1 for n in el.iter() if n.tag in (W+'drawing',W+'pict'))
   blocks.append({'location':where,'text':t,'links':links,'imageContainers':images})
  def walk(parent,where):
   for i,x in enumerate(parent,1):
    loc=where+'/'+str(i)
    if x.tag==W+'p':paragraph(x,loc)
    elif x.tag==W+'tbl':
     for ri,row in enumerate(x.findall(W+'tr'),1):
      for ci,cell in enumerate(row.findall(W+'tc'),1):walk(cell,loc+f'/table-row-{ri}/cell-{ci}')
    elif x.tag in (W+'sdt',W+'sdtContent'):walk(x,loc)
  if body is None:raise ValueError('Missing document body')
  walk(body,'body')
  return {'sourceFile':Path(source).name,'blocks':blocks,'limitations':['Main document body only; headers, footers, comments, tracked-deletion text and metadata require separate review.','Images are counted, not interpreted. Table cell relationships/layout may require visual inspection.','This output has not been anonymised.']}
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('output');a=p.parse_args();data=extract(a.source)
 with Path(a.output).open('x',encoding='utf-8') as f:json.dump(data,f,ensure_ascii=False,indent=2);f.write('\n')
 print(json.dumps({'blocks':len(data['blocks']),'output':a.output}))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,KeyError,zipfile.BadZipFile,ET.ParseError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
