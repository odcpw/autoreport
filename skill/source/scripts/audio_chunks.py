#!/usr/bin/env python3
"""Split only when needed; keep original-time offsets. Requires ffmpeg/ffprobe."""
import argparse,hashlib,json,math,shutil,subprocess,sys
from pathlib import Path

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('output_dir');p.add_argument('--minutes',type=float,default=20);p.add_argument('--overlap',type=float,default=5);args=p.parse_args()
 if not shutil.which('ffmpeg') or not shutil.which('ffprobe'):raise ValueError('ffmpeg and ffprobe must be available')
 size=args.minutes*60
 if not math.isfinite(size) or not math.isfinite(args.overlap) or size<=0 or not 0<=args.overlap<size:raise ValueError('Require positive chunk length and 0 <= overlap < chunk length')
 src=Path(args.source).resolve();out=Path(args.output_dir)
 probe=subprocess.run(['ffprobe','-v','error','-show_entries','format=duration','-of','json',str(src)],capture_output=True,text=True,check=True)
 duration=float(json.loads(probe.stdout)['format']['duration'])
 if not math.isfinite(duration) or duration<=0:raise ValueError('Unknown/nonpositive duration')
 out.mkdir(parents=True,exist_ok=False)
 digest=hashlib.sha256()
 with src.open('rb') as f:
  for block in iter(lambda:f.read(1024*1024),b''):digest.update(block)
 manifest={'sourceFile':src.name,'sourceSha256':digest.hexdigest(),'sourceDurationSeconds':duration,'overlapSeconds':args.overlap,'complete':False,'chunks':[],'note':'Chunk transcript timestamps + sourceStartSeconds = original recording time. Deduplicate overlap; this script does not transcribe.'}
 mp=out/'audio_manifest.json'
 def save():mp.write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
 save();start=0.0;index=1
 while start<duration:
  end=min(start+size,duration);name=f'chunk_{index:04d}.mp3';target=out/name
  subprocess.run(['ffmpeg','-nostdin','-v','error','-n','-ss',str(start),'-i',str(src),'-t',str(end-start),'-map','0:a:0','-vn','-ac','1','-ar','16000','-c:a','libmp3lame','-b:a','64k','-map_metadata','-1',str(target)],check=True)
  if not target.stat().st_size:raise ValueError('Empty audio chunk')
  manifest['chunks'].append({'file':name,'sourceStartSeconds':round(start,6),'sourceEndSeconds':round(end,6),'bytes':target.stat().st_size,'transcriptionStatus':'not_started'});save()
  if end>=duration:break
  start=end-args.overlap;index+=1
 manifest['complete']=True;save();print(json.dumps({'manifest':str(mp),'chunks':len(manifest['chunks']),'sourceUnchanged':True}))
if __name__=='__main__':
 try:main()
 except (ValueError,OSError,KeyError,subprocess.CalledProcessError) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(1)
