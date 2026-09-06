#!/usr/bin/env python3
"""Narrow copy-only editor for the nested AutoBericht sidecar; no AI or network."""
import argparse,copy,hashlib,json,sys
from pathlib import Path

WS_FIELDS={'findingText','recommendationText','selectedLevel','includeFinding','includeRecommendation','done','priority'}
DERIVED={'scoreTouched','autoScoreLevel','findingLibraryAction','libraryAction'}
GROUPS=('report','observations','training')

def fail(message):raise ValueError(message)
def unique_pairs(pairs):
 d={}
 for k,v in pairs:
  if k in d:fail('Duplicate JSON key: '+k)
  d[k]=v
 return d
def read(path):
 data=Path(path).read_bytes()
 return json.loads(data.decode('utf-8-sig'),object_pairs_hook=unique_pairs),hashlib.sha256(data).hexdigest()
def write_new(path,doc):
 p=Path(path)
 with p.open('x',encoding='utf-8') as f:json.dump(doc,f,ensure_ascii=False,indent=2);f.write('\n')
def project(doc):
 p=doc.get('report',{}).get('project')
 if not isinstance(p,dict) or not isinstance(p.get('chapters'),list):fail('Expected report.project.chapters; legacy/unknown schema needs adaptation')
 return p
def row_map(doc):
 out={};chapters=set()
 for c in project(doc)['chapters']:
  cid=str(c.get('id',''))
  if not cid or cid in chapters:fail('Missing or duplicate chapter ID')
  chapters.add(cid)
  if not isinstance(c.get('rows'),list):fail('Chapter rows must be a list')
  for r in c['rows']:
   if r.get('kind')=='section':continue
   rid=str(r.get('id',''));key=(cid,rid)
   if not rid or key in out:fail('Missing or duplicate row ID within chapter')
   out[key]=r
 return out
def options(doc,group):
 opts=doc.get('photos',{}).get('photoTagOptions',{}).get(group,[])
 return {str(x.get('value','')) if isinstance(x,dict) else str(x) for x in opts}
def check_ws(row):
 ws=row.get('workstate',{})
 if not isinstance(ws,dict):fail('workstate must be an object')
 for key in ['includeFinding','includeRecommendation','done','scoreTouched']:
  if key in ws and type(ws[key]) is not bool:fail(key+' must be boolean')
 if 'selectedLevel' in ws and (type(ws['selectedLevel']) is not int or ws['selectedLevel'] not in range(1,5)):fail('selectedLevel must be an integer 1..4')
 if 'priority' in ws and (type(ws['priority']) is not int or ws['priority'] not in range(5)):fail('priority must be an integer 0..4')
 for key in ['findingText','recommendationText']:
  if key in ws and not isinstance(ws[key],str):fail(key+' must be a string')
def validate(before,after):
 oldrows=row_map(before);newrows=row_map(after)
 if oldrows.keys()!=newrows.keys():fail('Row identities changed')
 stripped=copy.deepcopy(after);strippedrows=row_map(stripped);changed=[]
 for key,new in newrows.items():
  old=oldrows[key];check_ws(new)
  if new==old:continue
  if old.get('kind')=='section':fail('Section rows cannot be edited')
  ow=old.get('workstate',{});nw=new.get('workstate',{})
  deltas={k for k in set(ow)|set(nw) if ow.get(k)!=nw.get(k) or (k in ow)!=(k in nw)}
  if deltas-(WS_FIELDS|DERIVED):fail('Unsupported workstate change: '+str(deltas))
  for k in ('findingLibraryAction','libraryAction'):
   if k in deltas and nw.get(k)!='off':fail('Draft editor cannot queue library actions')
  if 'selectedLevel' in deltas:
   if old.get('type') in ('field_observation','summary'):fail('No score for observation/summary')
   if nw.get('scoreTouched') is not True:fail('Changed assessment requires scoreTouched=true')
  copyrow=strippedrows[key]
  if 'workstate' in old:copyrow['workstate']=copy.deepcopy(ow)
  else:copyrow.pop('workstate',None)
  if copyrow!=old:fail('Row data outside workstate changed')
  changed.append({'chapterId':key[0],'rowId':key[1],'fields':sorted(deltas)})
 oldphotos=before.get('photos',{}).get('photos',{})
 newphotos=after.get('photos',{}).get('photos',{})
 photo_changes=[]
 for path,photo in newphotos.items():
  old=oldphotos.get(path)
  if photo==old:continue
  if not isinstance(photo,dict):fail('Photo entry must be an object')
  for k in set(photo)|(set(old) if old else set()):
   if k not in ('tags','notes') and (old is None or photo.get(k)!=old.get(k)):fail('Unsupported photo field change')
  if old is None and set(photo)-{'tags','notes'}:fail('New photo entries may contain only tags and notes')
  if 'notes' in photo and not isinstance(photo['notes'],str):fail('Photo notes must be text')
  tags=photo.get('tags',{});oldtags=(old or {}).get('tags',{})
  if set(tags)-set(GROUPS):fail('Unknown photo tag group')
  for group in GROUPS:
   vals=tags.get(group,[])
   if not isinstance(vals,list) or any(not isinstance(v,str) for v in vals):fail('Photo tags must be string lists')
   if len(vals)!=len(set(vals)):fail('Duplicate photo tags')
   if set(vals)-set(oldtags.get(group,[]))-options(before,group):fail('Unknown new tag value in '+group)
  photo_changes.append(path)
 if set(oldphotos)-set(newphotos):fail('Photo records removed')
 if photo_changes:
  if 'photos' not in before:fail('Photo branch absent; initialize through app first')
  stripped['photos']['photos']=copy.deepcopy(oldphotos)
 if stripped!=before:fail('Unrelated sidecar data changed')
 return {'valid':True,'rowChanges':changed,'photoChanges':photo_changes,'validationLevel':'JSON contract and preservation; not app/UI import'}
def apply(doc,plan,sha):
 if plan.get('format')!='autobericht-editor-patch/1':fail('Unknown patch format')
 if plan.get('sourceSha256')!=sha:fail('Source hash mismatch; regenerate against the latest sidecar')
 if set(plan)-{'format','sourceSha256','rows','photos'}:fail('Unsupported patch keys')
 out=copy.deepcopy(doc);rows=row_map(out);seen=set()
 for edit in plan.get('rows',[]):
  if set(edit)!={'chapterId','rowId','changes'}:fail('Row edit keys must be chapterId,rowId,changes')
  key=(str(edit['chapterId']),str(edit['rowId']))
  if key in seen or key not in rows:fail('Duplicate or unknown row target '+str(key))
  seen.add(key);row=rows[key]
  if row.get('kind')=='section':fail('Cannot edit section row')
  changes=edit['changes']
  if not isinstance(changes,dict) or set(changes)-WS_FIELDS:fail('Unsupported row edit field')
  ws=row.setdefault('workstate',{});material=any(ws.get(k)!=v for k,v in changes.items() if k!='done')
  if 'selectedLevel' in changes:
   if row.get('type') in ('field_observation','summary'):fail('No score for observation/summary')
   ws['scoreTouched']=True;ws['autoScoreLevel']=changes['selectedLevel']
  if 'findingText' in changes:ws['findingLibraryAction']='off'
  if 'recommendationText' in changes:ws['libraryAction']='off'
  ws.update(changes)
  if material and 'done' not in changes:ws['done']=False
  check_ws(row)
 seenphotos=set()
 for edit in plan.get('photos',[]):
  allowed={'path','notes'}|{g+s for g in GROUPS for s in ('Add','Remove')}
  if set(edit)-allowed:fail('Unsupported photo edit key')
  path=edit.get('path')
  if not isinstance(path,str) or not path or path in seenphotos:fail('Missing or duplicate photo path')
  if path.startswith(('/', '\\')) or '..' in path.replace('\\','/').split('/'):fail('Photo path must be project-relative')
  seenphotos.add(path)
  if 'photos' not in out:fail('Initialize photo branch through app first')
  ph=out['photos'].setdefault('photos',{}).setdefault(path,{'notes':'','tags':{g:[] for g in GROUPS}})
  tags=ph.setdefault('tags',{})
  for g in GROUPS:
   if g+'Add' not in edit and g+'Remove' not in edit:continue
   adds=edit.get(g+'Add',[]);removes=edit.get(g+'Remove',[])
   if not isinstance(adds,list) or not isinstance(removes,list) or any(not isinstance(x,str) for x in adds+removes):fail('Tag changes must be string lists')
   if set(adds)&set(removes):fail('Same tag added and removed')
   if set(adds)-options(doc,g):fail('Unknown tag value in '+g)
   values=[x for x in tags.get(g,[]) if x not in removes]
   tags[g]=list(dict.fromkeys(values+adds))
  if 'notes' in edit:ph['notes']=edit['notes']
 validate(doc,out)
 return out

def main():
 parser=argparse.ArgumentParser(description=__doc__);sub=parser.add_subparsers(dest='command',required=True)
 p=sub.add_parser('inspect');p.add_argument('source')
 p=sub.add_parser('apply');p.add_argument('source');p.add_argument('plan');p.add_argument('output')
 p=sub.add_parser('validate');p.add_argument('source');p.add_argument('result')
 args=parser.parse_args();doc,sha=read(args.source)
 if args.command=='inspect':
  rows=row_map(doc)
  print(json.dumps({'sourceSha256':sha,'locale':project(doc).get('meta',{}).get('locale'),'rows':[{'chapterId':key[0],'rowId':key[1],**row} for key,row in rows.items()],'photoTagOptions':doc.get('photos',{}).get('photoTagOptions',{}),'photoPaths':list(doc.get('photos',{}).get('photos',{}))},ensure_ascii=False,indent=2))
 elif args.command=='apply':
  plan,_=read(args.plan);out=apply(doc,plan,sha);report=validate(doc,out);write_new(args.output,out);print(json.dumps(report,ensure_ascii=False,indent=2))
 else:
  result,_=read(args.result);print(json.dumps(validate(doc,result),ensure_ascii=False,indent=2))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,TypeError,KeyError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
