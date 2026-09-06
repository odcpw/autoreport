#!/usr/bin/env python3
"""Synthetic behaviour checks for the portable helpers, without client data."""
import copy,hashlib,json,shutil,subprocess,sys,tempfile,unittest,zipfile
from pathlib import Path
import sidecar_tool as sc
import library_tool as lib
import extract_docx as dx

def fixture():
 return {'report':{'project':{'meta':{'locale':'fr-CH'},'chapters':[{'id':'1','rows':[{'id':'1.1','type':'standard','master':{'finding':'Generic','recommendation':'Base'},'customer':{'answer':1,'items':[{'id':'1.1','comment':'Customer statement','answer':1}]},'workstate':{'selectedLevel':4,'scoreTouched':False,'findingText':'Generic','recommendationText':'Base','done':True,'includeFinding':True,'libraryAction':'append'}}]},{'id':'4.8','rows':[{'id':'4.8.1','type':'field_observation','tag':'Rayonnages','master':{'finding':'Observation','recommendation':'Base observation'},'customer':{},'workstate':{'selectedLevel':1,'done':False}}]}]}},'photos':{'photoRoot':'photos','photoTagOptions':{'observations':[{'value':'Rayonnages','label':'Rayonnages'}],'report':['1.1'],'training':['Training']},'photos':{'photos/a.jpg':{'tags':{'report':['1.1'],'observations':[],'training':['Training']},'notes':'Keep me'}}},'spider':{'unrelated':[1,2]},'unknownFutureField':{'keep':True}}

class Helpers(unittest.TestCase):
 def test_case_edit_preserves_source_and_review_gate(self):
  source=fixture();source['report']['project']['chapters'][0]['rows'].append({'kind':'section','title':'Section heading without row ID'});snapshot=copy.deepcopy(source)
  plan={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[{'chapterId':'1','rowId':'1.1','changes':{'selectedLevel':2,'recommendationText':'A case-specific draft'}}],'photos':[{'path':'photos/a.jpg','observationsAdd':['Rayonnages']}]}
  out=sc.apply(source,plan,'test');self.assertEqual(source,snapshot)
  r=out['report']['project']['chapters'][0]['rows'][0]
  self.assertFalse(r['workstate']['done']);self.assertTrue(r['workstate']['scoreTouched']);self.assertEqual(r['workstate']['libraryAction'],'off')
  self.assertEqual(r['customer'],source['report']['project']['chapters'][0]['rows'][0]['customer'])
  self.assertEqual(r['master'],source['report']['project']['chapters'][0]['rows'][0]['master'])
  self.assertEqual(out['photos']['photos']['photos/a.jpg']['notes'],'Keep me');self.assertEqual(out['photos']['photos']['photos/a.jpg']['tags']['training'],['Training'])
  self.assertEqual(out['spider'],source['spider']);self.assertTrue(sc.validate(source,json.loads(json.dumps(out)))['valid'])
 def test_rejects_stale_unknown_invalid_and_unrelated_changes(self):
  source=fixture();base={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[]}
  with self.assertRaises(ValueError):sc.apply(source,base,'different')
  for change in [{'selectedLevel':80},{'selectedLevel':True},{'priority':7},{'invented':1}]:
   p={**base,'rows':[{'chapterId':'1','rowId':'1.1','changes':change}]}
   with self.assertRaises(ValueError):sc.apply(source,p,'test')
  p={**base,'photos':[{'path':'photos/a.jpg','observationsAdd':['Unknown category']} ]}
  with self.assertRaises(ValueError):sc.apply(source,p,'test')
  p={**base,'rows':[{'chapterId':'4.8','rowId':'4.8.1','changes':{'selectedLevel':2}}]}
  with self.assertRaises(ValueError):sc.apply(source,p,'test')
  changed=copy.deepcopy(source);changed['report']['project']['chapters'][0]['rows'][0]['customer']['answer']=0
  with self.assertRaises(ValueError):sc.validate(source,changed)
  with self.assertRaises(ValueError):sc.row_map({'chapters':[]})
 def test_explicit_done_and_no_overwrite(self):
  source=fixture();plan={'format':'autobericht-editor-patch/1','sourceSha256':'test','rows':[{'chapterId':'1','rowId':'1.1','changes':{'findingText':'Reviewed new text','done':True}}]}
  out=sc.apply(source,plan,'test');self.assertTrue(out['report']['project']['chapters'][0]['rows'][0]['workstate']['done'])
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'out.json';sc.write_new(p,out)
   with self.assertRaises(FileExistsError):sc.write_new(p,out)
 def test_library_addition_and_repeat(self):
  source={'schemaVersion':1,'structure':{'items':[]},'library':{'entries':[{'id':'1.1','finding':'Generic','recommendation':'First'}],'observations':[]},'tags':{'keep':True}}
  item={'kind':'entry','key':'1.1','text':'Second useful variant','reviewed':True,'generalised':True}
  plan={'format':'library-additions/1','sourceSha256':'test','additions':[item,item]}
  out,report=lib.merge(source,plan,'test');self.assertEqual(report['applied'],['1.1']);self.assertEqual(report['exactDuplicatesSkipped'],['1.1']);self.assertEqual(out['library']['entries'][0]['finding'],'Generic')
  self.assertEqual(source['library']['entries'][0]['recommendation'],'First')
  _,repeat=lib.merge(out,plan,'test');self.assertEqual(repeat['applied'],[])
  with self.assertRaises(ValueError):lib.merge(source,{**plan,'additions':[{**item,'reviewed':False}]},'test')
 def test_docx_table_and_hyperlink(self):
  xml='''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Introduction</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Finding</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Recommendation </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>reference</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'synthetic.docx'
   with zipfile.ZipFile(p,'w') as z:
    z.writestr('word/document.xml',xml);z.writestr('word/_rels/document.xml.rels','<Relationships><Relationship Id="rId1" Target="https://example.org/reference"/></Relationships>')
   data=dx.extract(p);self.assertEqual([b['text'] for b in data['blocks']],['Introduction','Finding','Recommendation reference']);self.assertIn('/cell-2/',data['blocks'][2]['location']);self.assertEqual(data['blocks'][2]['links'][0]['target'],'https://example.org/reference')
 @unittest.skipUnless(shutil.which('ffmpeg') and shutil.which('ffprobe'),'ffmpeg/ffprobe unavailable')
 def test_audio_chunk_offsets_and_original(self):
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);src=root/'synthetic.wav'
   subprocess.run(['ffmpeg','-v','error','-f','lavfi','-i','sine=frequency=400:duration=4','-c:a','pcm_s16le',str(src)],check=True)
   original=hashlib.sha256(src.read_bytes()).hexdigest()
   subprocess.run([sys.executable,str(Path(__file__).with_name('audio_chunks.py')),str(src),str(root/'chunks'),'--minutes','0.03','--overlap','0.2'],capture_output=True,text=True,check=True)
   m=json.loads((root/'chunks/audio_manifest.json').read_text());self.assertTrue(m['complete']);self.assertEqual(m['sourceSha256'],original);self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(),original)
   self.assertEqual(m['chunks'][0]['sourceStartSeconds'],0);self.assertEqual(m['chunks'][-1]['sourceEndSeconds'],4)
   for a,b in zip(m['chunks'],m['chunks'][1:]):self.assertAlmostEqual(a['sourceEndSeconds']-b['sourceStartSeconds'],0.2)
if __name__=='__main__':unittest.main(verbosity=2)
