import {loadCachedSnapshot} from './d1-calendar.js';
export const DOCTOR_CALENDAR_VERSION='doctor-month-v1';
export const DOCTOR_CALENDAR_MAX_BYTES=2*1024*1024;
const sources=new Set(['mmc','mch','ddh','vhh','casey']);
export function validDoctorMonth(source,month,revision){return sources.has(source)&&/^\d{4}-(0[1-9]|1[0-2])$/.test(month)&&/^[a-f0-9]{64}$/.test(revision);}
export function doctorMonthIndexKey(source,month,revision){if(!validDoctorMonth(source,month,revision))throw Error('Invalid doctor calendar generation.');return `doctor-calendars/v1/${source}/${month}/indexes/${revision}.json`;}
export function doctorMonthRowsKey(source,month,digest){if(!validDoctorMonth(source,month,digest))throw Error('Invalid doctor calendar content.');return `doctor-calendars/v1/${source}/${month}/rows/${digest}.json`;}
export async function loadDoctorMonthSelection(r2,source,month,revision,keys,{metadataOnly=false}={}){
 const index=await loadCachedSnapshot(r2,doctorMonthIndexKey(source,month,revision));
 if(!index||index.version!==DOCTOR_CALENDAR_VERSION||index.sourceType!==source||index.month!==month||index.parentRevision!==revision||!index.doctors||Object.keys(index.doctors).length>1024)return null;
 const rows=[],selected=[];
 for(const doctorKey of [...new Set(keys)].sort()){
  const pointer=Object.hasOwn(index.doctors,doctorKey)?index.doctors[doctorKey]:null;
  if(!pointer){selected.push([doctorKey,'absent']);continue;}
  if(!/^[a-f0-9]{64}$/.test(pointer.digest)||pointer.key!==doctorMonthRowsKey(source,month,pointer.digest))return null;
  selected.push([doctorKey,pointer.digest]);
  if(metadataOnly)continue;
  const payload=await loadCachedSnapshot(r2,pointer.key);
  if(!payload||payload.version!==DOCTOR_CALENDAR_VERSION||payload.sourceType!==source||payload.month!==month||payload.doctorKey!==doctorKey||!Array.isArray(payload.rows)||payload.rows.length>2048||payload.rows.some(r=>r.sourceType!==source||r.doctorKey!==doctorKey))return null;
  rows.push(...payload.rows);
 }
 return {rows,selected};
}
