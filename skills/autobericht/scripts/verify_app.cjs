#!/usr/bin/env node
// Optional validation against a supplied AutoBericht checkout. No network.
const fs=require('node:fs');const path=require('node:path');const vm=require('node:vm');const assert=require('node:assert/strict');
const [repo,libraryPath,sidecarPath]=process.argv.slice(2);
if(!repo||!libraryPath){console.error('Usage: node verify_app.cjs REPO LIBRARY_JSON [SIDECAR_JSON]');process.exit(2);}
const hooks={};const ctx={console,structuredClone,TextEncoder,TextDecoder,URL,setTimeout,clearTimeout,__AUTO_BERICHT_TEST__:hooks,document:{getElementById(){return null;},addEventListener(){},visibilityState:'visible'},addEventListener(){},AutoReportDebug:{logLine(){}},AutoBerichtI18n:{t:(k,f)=>f||k,tf:(k,f)=>f||k,setLocale(){},resolveSpellcheckLang:l=>l||'fr-CH'},AutoBerichtFsHandle:{saveHandle:async()=>{},loadHandle:async()=>null,requestHandlePermission:async()=>false}};ctx.window=ctx;ctx.globalThis=ctx;vm.createContext(ctx);
for(const file of ['dependencies.js','state.js','normalize.js','seeds.js','report-rows.js']){const p=path.join(repo,'AutoBericht/mini/shared',file);vm.runInContext(fs.readFileSync(p,'utf8'),ctx,{filename:p});}
const makerPath=path.join(repo,'AutoBericht/mini/librarymaker.js');vm.runInContext(fs.readFileSync(makerPath,'utf8'),ctx,{filename:makerPath});
const library=JSON.parse(fs.readFileSync(libraryPath,'utf8'));ctx.AutoBerichtSeeds.validateKnowledgeBase(library);
const maker=hooks.librarymaker?.normalizeKnowledgeBaseForMaker;assert.equal(typeof maker,'function','LibraryMaker validation hook unavailable in this revision');
const normal=maker(library);
for(const group of ['entries','observations']){const key=group==='entries'?'id':'value';for(const item of library.library[group]||[]){const actual=normal.library[group].find(x=>x[key]===item[key]);assert(actual);assert.equal(actual.finding,item.finding||'');assert.equal(actual.recommendation,item.recommendation||'');}}
const project=ctx.AutoBerichtSeeds.buildProjectFromKnowledgeBase(library);const text=JSON.stringify(project);
// The importer addresses standard entries by collapsedId/groupId, not every raw
// questionnaire sub-item ID. A schema-valid library can therefore lose text at
// project creation. Report coverage separately and fail rather than hiding loss.
const missingRecommendations=[];
for(const item of [...(library.library.entries||[]),...(library.library.observations||[])]){if(item.recommendation&&!text.includes(JSON.stringify(item.recommendation).slice(1,-1)))missingRecommendations.push(item.id||item.value);}
let checked=0;
if(sidecarPath){
 const sidecar=JSON.parse(fs.readFileSync(sidecarPath,'utf8'));assert(sidecar.report?.project?.chapters,'Expected nested sidecar');
 const copy=JSON.parse(JSON.stringify(sidecar));ctx.AutoBerichtNormalize.normalizeProject(copy.report.project);ctx.AutoBerichtNormalize.syncObservationChapterRows(copy.report.project,copy);
 for(const c of sidecar.report.project.chapters){for(const row of c.rows||[]){if(row.kind==='section')continue;const actual=copy.report.project.chapters.find(x=>String(x.id)===String(c.id))?.rows.find(x=>x.kind!=='section'&&x.id===row.id);assert(actual,`Row lost after normalisation: ${row.id}`);assert.deepEqual(actual.customer,row.customer);
  for(const key of ['findingText','recommendationText','selectedLevel','includeFinding','includeRecommendation','done']){if(Object.hasOwn(row.workstate||{},key))assert.equal(actual.workstate[key],row.workstate[key],`Changed ${row.id}/${key}`);}
  if(row.type==='field_observation'||row.type==='summary')assert.equal(ctx.AutoBerichtState.calculateScore(actual),null);checked++;
 }}
 assert.deepEqual(copy.photos,sidecar.photos);
}
console.log(JSON.stringify({librarySchema:'passed',libraryMaker:'passed',projectImport:'passed',recommendationCoverage:missingRecommendations.length?'failed':'passed',missingRecommendations,sidecarRowsChecked:checked,interactiveImportTested:false},null,2));
if(missingRecommendations.length)process.exitCode=1;
