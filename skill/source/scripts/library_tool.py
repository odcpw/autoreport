#!/usr/bin/env python3
"""Append reviewed recommendation variants to an existing library copy."""
import argparse,copy,hashlib,json,re,sys
from pathlib import Path

def normal(text):return re.sub(r'\s+',' ',text).strip()
def merge(source,plan,sha):
 if plan.get('format')!='library-additions/1' or plan.get('sourceSha256')!=sha:raise ValueError('Wrong format or stale source hash')
 if set(plan)-{'format','sourceSha256','additions'}:raise ValueError('Unsupported plan field')
 if not isinstance(source.get('library'),dict):raise ValueError('Missing library')
 out=copy.deepcopy(source);applied=[];skipped=[]
 for item in plan.get('additions',[]):
  if set(item)!={'kind','key','text','reviewed','generalised'}:raise ValueError('Addition requires kind,key,text,reviewed,generalised')
  if item['reviewed'] is not True or item['generalised'] is not True:raise ValueError('Only reviewed, generalised text may be merged')
  if item['kind'] not in ('entry','observation'):raise ValueError('Unknown target kind')
  group='entries' if item['kind']=='entry' else 'observations';field='id' if group=='entries' else 'value'
  found=[x for x in out['library'].get(group,[]) if str(x.get(field))==str(item['key'])]
  if len(found)!=1:raise ValueError('Target must already exist uniquely')
  text=item['text']
  if not isinstance(text,str) or not text.strip():raise ValueError('Non-empty recommendation text required')
  current=found[0].get('recommendation','')
  if not isinstance(current,str):raise ValueError('Expected recommendation string')
  blocks=re.split(r'\n\s*---\s*\n',current)
  if normal(text) in {normal(x) for x in blocks}:
   skipped.append(item['key']);continue
  found[0]['recommendation']=(current.rstrip()+'\n\n---\n\n' if current.strip() else '')+text.strip()
  applied.append(item['key'])
 # Prove all data other than the targeted recommendation strings stayed intact.
 stripped=copy.deepcopy(out)
 for group in ('entries','observations'):
  for old,new in zip(source['library'].get(group,[]),stripped['library'].get(group,[])):
   if 'recommendation' in old:new['recommendation']=old['recommendation']
   else:new.pop('recommendation',None)
 if stripped!=source:raise ValueError('Non-recommendation library data changed')
 return out,{'valid':True,'applied':applied,'exactDuplicatesSkipped':skipped,'genericFindingsPreserved':True,'semanticReviewPerformedByScript':False}
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('plan');p.add_argument('output');args=p.parse_args()
 raw=Path(args.source).read_bytes();source=json.loads(raw.decode('utf-8-sig'));plan=json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
 out,report=merge(source,plan,hashlib.sha256(raw).hexdigest())
 with Path(args.output).open('x',encoding='utf-8') as f:json.dump(out,f,ensure_ascii=False,indent=2);f.write('\n')
 print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,TypeError,KeyError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
