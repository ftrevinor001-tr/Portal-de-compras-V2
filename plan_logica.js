// Portal de Compras V2 · lógica de planeación compartida (Carga/cálculo y Tablero de planeación)
// No cambiar reglas aquí sin revisarlas con Compras: el cálculo diario y el tablero usan este mismo archivo.
/*PARSE-START*/
const P={};
P.norm=s=>String(s==null?'':s).trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ');
P.num=v=>{if(v==null||v==='')return null;if(typeof v==='number')return isFinite(v)?v:null;const n=parseFloat(String(v).replace(/[$,\s]/g,''));return isNaN(n)?null:n;};
P.clave=v=>{const n=P.num(v);return n!=null&&Number.isInteger(n)&&n>0?n:null;};
P.pad=n=>String(n).padStart(2,'0');
// Fecha de Excel (serial, texto DD/MM/AAAA o AAAA-MM-DD) → 'AAAA-MM-DD'
P.fecha=v=>{
  if(v==null||v==='')return null;
  if(typeof v==='number'){if(v<20000||v>80000)return null;const d=new Date(Math.round((v-25569)*864e5));return d.getUTCFullYear()+'-'+P.pad(d.getUTCMonth()+1)+'-'+P.pad(d.getUTCDate());}
  if(v instanceof Date&&!isNaN(v))return v.getFullYear()+'-'+P.pad(v.getMonth()+1)+'-'+P.pad(v.getDate());
  const s=String(v).trim();let m;
  if((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)))return m[1]+'-'+P.pad(m[2])+'-'+P.pad(m[3]);
  if((m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)))return m[3]+'-'+P.pad(m[2])+'-'+P.pad(m[1]);
  return null;
};
// Encabezado de mes → 'AAAA-MM-01'
P.mes=v=>{
  if(v==null||v==='')return null;
  if(typeof v==='number'){const f=P.fecha(v);return f?f.slice(0,7)+'-01':null;}
  const s=String(v).trim();let m;
  if((m=s.match(/^(\d{4})-(\d{1,2})$/)))return m[1]+'-'+P.pad(m[2])+'-01';
  const f=P.fecha(s);return f?f.slice(0,7)+'-01':null;
};
P.aoa=ws=>XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
// Busca la fila de encabezado (primeras 10) que contenga todas las columnas pedidas
P.header=(aoa,requeridas)=>{
  const req=requeridas.map(P.norm);
  for(let i=0;i<Math.min(10,aoa.length);i++){
    const fila=aoa[i].map(P.norm);
    if(req.every(r=>fila.includes(r))){const idx={};req.forEach((r,j)=>idx[requeridas[j]]=fila.indexOf(r));return{fila:i,idx,cols:fila};}
  }
  return null;
};
P.hoja=(wb,nombres)=>{for(const n of nombres){const k=wb.SheetNames.find(s=>P.norm(s)===P.norm(n));if(k)return wb.Sheets[k];}return null;};

// 1. Ventas: {empresas:{SANVER:{meses,rows},MF:{…}}, avisos}
P.ventas=wb=>{
  const out={empresas:{},avisos:[]};
  for(const [emp,hoja] of [['SANVER','VENTAS_SANVER'],['MF','VENTAS_MF']]){
    const ws=P.hoja(wb,[hoja]);
    if(!ws){out.avisos.push('No se encontró la hoja '+hoja+'.');continue;}
    const aoa=P.aoa(ws);
    const h=aoa[0]||[];
    if(P.norm(h[0])!=='CLAVE')out.avisos.push(hoja+': la columna A debería llamarse Clave.');
    const meses=[];h.forEach((c,j)=>{if(j===0)return;const m=P.mes(c);if(m)meses.push({j,m});});
    const acc=new Map();let neg=0,dup=0,malas=0;const vistas=new Set();
    for(let i=1;i<aoa.length;i++){
      const r=aoa[i];const k=P.clave(r[0]);
      if(k==null){if(String(r[0]).trim()!=='')malas++;continue;}
      if(vistas.has(k))dup++;vistas.add(k);
      for(const {j,m} of meses){
        const v=P.num(r[j]);if(v==null||v===0)continue;
        if(v<0)neg++;
        const key=k+'|'+m;acc.set(key,(acc.get(key)||0)+v);
      }
    }
    const porClave=new Map();let celdas=0;
    acc.forEach((u,key)=>{const [c,m]=key.split('|');if(u===0)return;const k=+c;
      if(!porClave.has(k))porClave.set(k,{});porClave.get(k)[m.slice(0,7)]=Math.round(u*10000)/10000;celdas++;});
    vistas.forEach(k=>{if(!porClave.has(k))porClave.set(k,{});});
    out.empresas[emp]={meses:meses.map(x=>x.m.slice(0,7)),claves:vistas.size,celdas,porClave};
    if(dup)out.avisos.push(hoja+': '+dup+' claves repetidas (se sumaron).');
    if(neg)out.avisos.push(hoja+': '+neg+' celdas negativas (devoluciones; se guardan tal cual).');
    if(malas)out.avisos.push(hoja+': '+malas+' filas con clave no numérica (se ignoraron).');
  }
  return out;
};

// Combina lo que ya está en la base con el archivo: los meses del archivo se REEMPLAZAN
// (en todas las claves); los demás meses se conservan. Devuelve solo las filas que cambian.
P.mergeVentas=(empresa,existentes,arch)=>{
  const meses=new Set(arch.meses);const out=[];
  const base=new Map(existentes.map(r=>[Number(r.clave),r.meses||{}]));
  const claves=new Set([...base.keys(),...arch.porClave.keys()]);
  claves.forEach(k=>{
    const antes=base.get(k)||{};const nuevo={};
    Object.keys(antes).forEach(m=>{if(!meses.has(m))nuevo[m]=antes[m];});
    Object.assign(nuevo,arch.porClave.get(k)||{});
    const ord={};Object.keys(nuevo).sort().forEach(m=>ord[m]=nuevo[m]);
    if(JSON.stringify(ord)!==JSON.stringify(antes)||!base.has(k))out.push({empresa,clave:k,meses:ord});
  });
  return out;
};

// 2. Existencias MF: {fecha, rows:[{clave,existencia}], sucursales, negativas}
P.existMF=wb=>{
  let ws=P.hoja(wb,['Existencias'])||wb.Sheets[wb.SheetNames[0]];
  const aoa=P.aoa(ws);const h=P.header(aoa,['CLAVE','EXISTENCIA']);
  if(!h)throw new Error('No encontré las columnas CLAVE y EXISTENCIA en la hoja '+(wb.SheetNames[0]));
  const iF=h.cols.indexOf('FECHA'),iS=h.cols.indexOf('SUCURSAL');
  const sum=new Map(),suc=new Set();let fmax=null;
  for(let i=h.fila+1;i<aoa.length;i++){
    const r=aoa[i];const k=P.clave(r[h.idx.CLAVE]);if(k==null)continue;
    sum.set(k,(sum.get(k)||0)+(P.num(r[h.idx.EXISTENCIA])||0));
    if(iS>=0&&r[iS]!=='')suc.add(String(r[iS]).trim());
    if(iF>=0){const f=P.fecha(r[iF]);if(f&&(!fmax||f>fmax))fmax=f;}
  }
  let neg=0;const rows=[];
  sum.forEach((e,k)=>{if(e<0)neg++;rows.push({clave:k,existencia:Math.max(0,Math.round(e*10000)/10000)});});
  return{fecha:fmax,rows,sucursales:suc.size,negativas:neg};
};

// 3. Empaques: rows [{clave,uni_compra,inner_cant,master_cant,factor_conv,validacion}]
P.empaques=wb=>{
  const ws=P.hoja(wb,['BD']);if(!ws)throw new Error('No se encontró la hoja BD.');
  const aoa=P.aoa(ws);const h=P.header(aoa,['CLAVE SANVER','UNI COMPRA','CTD. EN UM','CTD. EN UM2','VALIDACION']);
  if(!h)throw new Error('No encontré los encabezados de la hoja BD (CLAVE SANVER, UNI COMPRA, Ctd. En UM, Ctd. En UM2, VALIDACION).');
  const iC=h.cols.indexOf('CAMBIO DE UNIDADES');
  const m=new Map();
  for(let i=h.fila+1;i<aoa.length;i++){
    const r=aoa[i];const k=P.clave(r[h.idx['CLAVE SANVER']]);if(k==null)continue;
    const fila={clave:k,uni_compra:String(r[h.idx['UNI COMPRA']]||'').trim()||null,
      inner_cant:P.num(r[h.idx['CTD. EN UM']]),master_cant:P.num(r[h.idx['CTD. EN UM2']]),
      factor_conv:iC>=0?P.num(r[iC]):null,validacion:P.norm(r[h.idx.VALIDACION])||null};
    const prev=m.get(k);
    if(!prev||(prev.validacion!=='VALIDO'&&fila.validacion==='VALIDO'))m.set(k,fila);
  }
  const rows=[...m.values()];
  return{rows,validos:rows.filter(r=>r.validacion==='VALIDO').length};
};

// 4. Histórico de agotados SANVER: rows [{empresa,clave,fecha,dias}]
P.historico=wb=>{
  const set=new Map();let leidas=0,dias=new Set();
  const leer=(ws,colsFecha)=>{
    if(!ws)return false;
    const aoa=P.aoa(ws);
    let h=null;for(const cf of colsFecha){h=P.header(aoa,[cf,'CLAVE','EXISTENCIAS']);if(h){h.cf=cf;break;}}
    if(!h)return false;
    for(let i=h.fila+1;i<aoa.length;i++){
      const r=aoa[i];const k=P.clave(r[h.idx.CLAVE]);if(k==null)continue;leidas++;
      const f=P.fecha(r[h.idx[h.cf]]);if(!f)continue;dias.add(f);
      const e=P.num(r[h.idx.EXISTENCIAS]);
      if(e!=null&&e<=0)set.set(k+'|'+f,{empresa:'SANVER',clave:k,fecha:f,dias:1});
    }
    return true;
  };
  const a=leer(P.hoja(wb,['Historico de faltantes']),['DIA/ MES.1','DIA/MES','FECHA']);
  const b=leer(P.hoja(wb,['Detalle']),['FECHA']);
  if(!a&&!b)throw new Error('No encontré las hojas Historico de faltantes ni Detalle con columnas Fecha, Clave y Existencias.');
  const ds=[...dias].sort();
  return{rows:[...set.values()],leidas,desde:ds[0],hasta:ds[ds.length-1],dias:ds.length};
};

// 5. Compras por entregas directas: rows [{clave,mes,cantidad}]
P.comprasED=wb=>{
  let ws=null,h=null;
  for(const n of wb.SheetNames){const aoa=P.aoa(wb.Sheets[n]);const hh=P.header(aoa,['MES/AÑO','CLAVE','TIPO DE SOLICITUD','CANT REC']);if(hh){ws=aoa;h=hh;break;}}
  if(!h)throw new Error('No encontré una hoja con las columnas Mes/año, CLAVE, TIPO DE SOLICITUD y CANT REC.');
  const acc=new Map(),meses=new Set();
  for(let i=h.fila+1;i<ws.length;i++){
    const r=ws[i];const m=P.mes(r[h.idx['MES/AÑO']]);if(!m)continue;meses.add(m);
    if(P.norm(r[h.idx['TIPO DE SOLICITUD']])!=='ENTREGAS DIRECTAS')continue;
    const k=P.clave(r[h.idx.CLAVE]);const q=P.num(r[h.idx['CANT REC']]);if(k==null||!q)continue;
    acc.set(k+'|'+m,(acc.get(k+'|'+m)||0)+q);
  }
  const rows=[];acc.forEach((q,key)=>{const [c,m]=key.split('|');rows.push({clave:+c,mes:m,cantidad:Math.round(q*10000)/10000});});
  return{rows,meses:[...meses].sort()};
};
// 5b. Costo unitario de compra por clave (para medir la afectación en dinero): TOTAL ÷ CANT REC
// de los últimos 3 meses con compra de cada clave (CEDI y entregas directas). Unidad = la de venta (CANT REC viene en unidad básica).
P.costosCompra=wb=>{
  let ws=null,h=null;
  for(const n of wb.SheetNames){const aoa=P.aoa(wb.Sheets[n]);const hh=P.header(aoa,['MES/AÑO','CLAVE','TOTAL','CANT REC']);if(hh){ws=aoa;h=hh;break;}}
  if(!h)return null;
  const acc=new Map();
  for(let i=h.fila+1;i<ws.length;i++){
    const r=ws[i];const m=P.mes(r[h.idx['MES/AÑO']]);const k=P.clave(r[h.idx.CLAVE]);const t=P.num(r[h.idx.TOTAL]),q=P.num(r[h.idx['CANT REC']]);
    if(!m||k==null||!(t>0)||!(q>0))continue;
    if(!acc.has(k))acc.set(k,new Map());const mm=acc.get(k);const x=mm.get(m)||{t:0,q:0};x.t+=t;x.q+=q;mm.set(m,x);
  }
  const rows=[];acc.forEach((mm,k)=>{const ms=[...mm.keys()].sort().slice(-3);let t=0,q=0;ms.forEach(m=>{t+=mm.get(m).t;q+=mm.get(m).q;});
    if(q>0)rows.push({clave:k,costo:Math.round(t/q*10000)/10000,mes:ms[ms.length-1],meses:ms.length});});
  return rows;
};
// 6. Detalle del maestro diario: rows por fecha
P.detalle=wb=>{
  const ws=P.hoja(wb,['Detalle']);if(!ws)throw new Error('No se encontró la hoja Detalle.');
  const aoa=P.aoa(ws);
  const req=['FECHA','CLAVE','EXISTENCIAS','COMPRADOR','CLASIFICACION POR MOVIMIENTO','TIEMPO DE ENTREGA','OBJETIVO_MESES DE INVENTARIO','BACK ORDER'];
  const h=P.header(aoa,req);if(!h)throw new Error('La hoja Detalle no trae las columnas esperadas: '+req.join(', '));
  const c=n=>h.cols.indexOf(n);
  const iN=c('NOMBRE'),iU=c('UNIDAD'),iG=c('GRUPO'),iE=c('CLASIFICACION_ESTATUS DE INVENTARIO'),iM=c('MARCA');
  const porFecha=new Map();
  for(let i=h.fila+1;i<aoa.length;i++){
    const r=aoa[i];const k=P.clave(r[h.idx.CLAVE]);const f=P.fecha(r[h.idx.FECHA]);if(k==null||!f)continue;
    if(!porFecha.has(f))porFecha.set(f,new Map());
    porFecha.get(f).set(k,{clave:k,nombre:iN>=0?String(r[iN]).trim():'',unidad:iU>=0?P.norm(r[iU]):'',grupo:iG>=0?String(r[iG]).trim():'',marca:iM>=0&&r[iM]!=null?String(r[iM]).trim():'',
      comprador:P.norm(r[h.idx.COMPRADOR]),rot:P.norm(r[h.idx['CLASIFICACION POR MOVIMIENTO']]),
      lt:P.num(r[h.idx['TIEMPO DE ENTREGA']]),pm:P.num(r[h.idx['OBJETIVO_MESES DE INVENTARIO']]),
      ex:P.num(r[h.idx.EXISTENCIAS])||0,bo:P.num(r[h.idx['BACK ORDER']])||0,estX:iE>=0?P.norm(r[iE]):''});
  }
  return porFecha;
};

// ── Reglas de cálculo (acordadas con Compras, oct-2026) ──
P.R={ED_DESDE:'2025-01',ED_HASTA:'2025-12',ANIO_LIMPIO:'2026-01',DIAS_AGOTADO:19,ATIPICO:3,MIN_MESES_MEDIANA:3,VENTANA_ATRAS:12,
     FACT_ENTREGA:1.4,FACT_MAX:1.4,EXCLUIR_ROT:['SIN COMPRA','NUEVOS']};
// Riesgo con margen (oct-2026): una clave ENTRA a riesgo con menos de T. entrega × 1.2 y solo SALE cuando pasa de T. entrega × 1.4.
// Entre 1.2 y 1.4 conserva el estatus del día anterior (evita que brinque de riesgo a óptimo y regrese por unas piezas).
// Sobre inventario: más de Objetivo × 1.4 (sin cambio). Los meses de inventario NO cuentan el BO.
P.R.FACT_ENTRA_RIESGO=1.2;
P.ym=(ym,k)=>{let[y,m]=ym.split('-').map(Number);m+=k;while(m>12){m-=12;y++;}while(m<1){m+=12;y--;}return y+'-'+P.pad(m);};
// Serie de ventas de una clave ya limpia de entregas directas (solo 2025, nunca negativa)
P.serie=(meses,ed)=>{
  const s=Object.assign({},meses||{});const notas=[];
  if(ed)Object.keys(ed).forEach(m=>{if(m>=P.R.ED_DESDE&&m<=P.R.ED_HASTA&&s[m]!=null&&s[m]>=ed[m]){s[m]=s[m]-ed[m];notas.push(m);}});
  return{s,ed:notas};
};
// Valor de un mes después de la regla de atípicos: null = descartado
P.valor=(s,m)=>{
  const x=Math.max(0,s[m]||0);
  const w=[];for(let i=1;i<=12;i++){const v=s[P.ym(m,-i)];if(v>0)w.push(v);}
  if(w.length<P.R.MIN_MESES_MEDIANA)return{v:x};
  w.sort((a,b)=>a-b);const med=w.length%2?w[(w.length-1)/2]:(w[w.length/2-1]+w[w.length/2])/2;
  if(x<=P.R.ATIPICO*med)return{v:x};
  return m<P.R.ANIO_LIMPIO?{v:null,at:'descartado'}:{v:P.R.ATIPICO*med,at:'topado'};
};
// ¿Mes de 2025 sospechoso de agotado? (sin datos de existencia antes de abr-2026)
// Clave regular (vendió ≥10 de los 12 meses previos) con una racha de ≥2 meses en cero.
P.R.AGOT_DESDE='2026-04';P.R.REGULAR=10;P.R.RACHA=2;P.R.INTERMITENTE=6;
P.sospechoso=(s,m,lo,hi)=>{
  if(m>=P.R.AGOT_DESDE||(s[m]||0)>0)return false;
  let n=0;for(let i=1;i<=12;i++){const x=P.ym(m,-i);if(lo&&x<lo)break;if((s[x]||0)>0)n++;}
  if(n<P.R.REGULAR)return false;
  let r=1;
  for(let x=P.ym(m,-1);(!lo||x>=lo)&&!((s[x]||0)>0);x=P.ym(x,-1))r++;
  for(let x=P.ym(m,1);(!hi||x<=hi)&&!((s[x]||0)>0);x=P.ym(x,1)){r++;if(r>40)break;}
  return r>=P.R.RACHA;
};
// Antecedente y pronóstico de una serie para el mes de referencia.
// Un mes NO cuenta si: estuvo agotado más de 18 días (DIAS_AGOTADO = 19; antes 10, cambiado oct-2026), es anterior a la primera venta de la clave,
// o es un mes de 2025 sospechoso de agotado. Si una ventana queda sin meses válidos → null
// (la demanda usa solo la otra; nunca se promedia contra un cero falso).
P.antPron=(s,ref,N,agot,ultimo,primero)=>{
  const notas=[];const fin=ultimo&&ultimo<P.ym(ref,-1)?ultimo:P.ym(ref,-1);
  if(fin!==P.ym(ref,-1))notas.push('ventas cargadas solo hasta '+fin);
  const conVenta=Object.keys(s).filter(m=>s[m]>0).sort();const ini=conVenta[0]||null;
  const hi=ultimo||fin;const ok=m=>ini!=null&&m>=ini&&!agot(m)&&!P.sospechoso(s,m,primero,hi);
  const a=[];for(let i=0;i<P.R.VENTANA_ATRAS&&a.length<N;i++){const m=P.ym(fin,-i);if(!ok(m))continue;const r=P.valor(s,m);if(r.at)notas.push(m+' '+r.at);if(r.v==null)continue;a.push(r.v);}
  const p=[];for(let i=0;i<N;i++){const m=P.ym(ref,i-12);if(ultimo&&m>ultimo)continue;if(!ok(m))continue;const r=P.valor(s,m);if(r.at)notas.push(m+' '+r.at);if(r.v==null)continue;p.push(r.v);}
  const avg=x=>x.length?x.reduce((t,v)=>t+v,0)/x.length:null;
  // Intermitente: en los 12 meses previos válidos vendió en ≤6 → la demanda usa el promedio de esos 12 meses
  const w=[];let nv=0;for(let i=1;i<=12;i++){const m=P.ym(ref,-i);if(ultimo&&m>ultimo)continue;if(!ok(m))continue;const r=P.valor(s,m);if(r.v==null)continue;w.push(r.v);if(r.v>0)nv++;}
  const inter=w.length>=6&&nv<=P.R.INTERMITENTE;
  if(ini==null)notas.push('sin ventas en el historial');
  else if(ini>P.ym(ref,-12))notas.push('primera venta '+ini);
  return{A:avg(a),P:avg(p),nA:a.length,nP:p.length,I:inter?avg(w):null,notas};
};
// Estatus por meses de inventario. prev = estatus del día anterior (para el margen entre 1.2 y 1.4)
P.clasificar=(mi,lt,pm,prev)=>{
  if(mi<lt*P.R.FACT_ENTRA_RIESGO)return'RIESGO';
  if(mi<lt*P.R.FACT_ENTREGA&&(prev==='RIESGO'||prev==='FALTANTE'))return'RIESGO';
  if(mi<=pm*P.R.FACT_MAX)return'OPTIMO';
  return'SOBRE INVENTARIO';
};
P.estatus=(rot,ex,D,lt,pm,prev)=>{
  if(rot==='SIN COMPRA')return'SIN COMPRA';
  if(ex<=0)return'FALTANTE';
  if(!D)return'RIESGO';
  return P.clasificar(ex/D,lt,pm,prev);
};
// Estatus del Excel recalculado con la misma regla, a partir de los meses de inventario del Detalle.
// orig = estatus que trae el Excel (SIN COMPRA se respeta). Meses vacíos o 0 con existencia = sin venta → RIESGO.
P.estatusExcel=(orig,ex,mi,lt,pm,prev)=>{
  if(orig==='SIN COMPRA')return'SIN COMPRA';
  if(!(ex>0))return'FALTANTE';
  if(mi==null||mi===''||!(Number(mi)>0))return'RIESGO';
  return P.clasificar(Number(mi),Number(lt)||0,Number(pm)||0,prev);
};
// Recalcula una serie de días: filas [{fecha,clave,estatus,existencias,meses_inv,t_entrega,obj_meses}] → Map 'fecha|clave' → estatus nuevo
P.cadenaExcel=filas=>{
  const porClave=new Map();filas.forEach(r=>{const k=String(r.clave);(porClave.get(k)||porClave.set(k,[]).get(k)).push(r);});
  const out=new Map();
  porClave.forEach((a,k)=>{a.sort((x,y)=>x.fecha<y.fecha?-1:x.fecha>y.fecha?1:0);let prev=null;
    a.forEach(r=>{const e=P.estatusExcel(String(r.estatus||'').toUpperCase(),Number(r.existencias)||0,r.meses_inv,r.t_entrega,r.obj_meses,prev);out.set(r.fecha+'|'+k,e);prev=e;});});
  return out;
};
// Redondeo a empaque: ≥1 MASTER → múltiplo de MASTER más cercano; menor → hacia arriba al INNER
P.empaque=(q,e,unidad)=>{
  if(!q||!e||e.validacion!=='VALIDO')return{q,ok:false};
  let f=1;if(P.norm(e.uni_compra)!==P.norm(unidad)){if(!e.factor_conv)return{q,ok:false,nota:'empaque en otra unidad sin factor'};f=e.factor_conv;}
  const inn=(e.inner_cant||0)*f,mas=(e.master_cant||0)*f;
  if(mas>0&&q>=mas)return{q:Math.max(mas,Math.round(q/mas)*mas),ok:true,tipo:'MASTER'};
  if(inn>0)return{q:Math.ceil(q/inn)*inn,ok:true,tipo:'INNER'};
  if(mas>0)return{q:mas,ok:true,tipo:'MASTER'};
  return{q,ok:false};
};
// Cálculo de un día completo. ctx: {ventas:{SANVER:Map,MF:Map}, ed:Map, agot:Set('EMP|clave|AAAA-MM'), invMF:Map, emp:Map, ultimo:{SANVER,MF}}
P.calcularDia=(fecha,filas,ctx)=>{
  const ref=fecha.slice(0,7);const out=[];
  filas.forEach(d=>{
    const k=d.clave;const lt=d.lt||0,pm=d.pm||0;const N=Math.max(1,Math.ceil(pm||1));
    const sv=P.serie(ctx.ventas.SANVER.get(k),ctx.ed.get(k));
    const ap=P.antPron(sv.s,ref,N,m=>ctx.agot.has('SANVER|'+k+'|'+m),ctx.ultimo.SANVER,ctx.primero&&ctx.primero.SANVER);
    const D=ap.I!=null?ap.I:(ap.A!=null&&ap.P!=null?(ap.A+ap.P)/2:(ap.A??ap.P??0));
    const pv=ctx.prevEst&&ctx.prevEst.get(k)||{};
    const est=P.estatus(d.rot,d.ex,D,lt,pm,pv.est);
    // BO del CEDI: se quita la parte que es de OCs de entrega directa a sucursal (no llega al CEDI)
    const fr=ctx.boFrac&&ctx.boFrac.has(k)?ctx.boFrac.get(k):1;const bo=Math.round(d.bo*fr*10000)/10000;
    const sug=X=>X==null||P.R.EXCLUIR_ROT.includes(d.rot)?null:Math.max(0,Math.ceil(X*(lt+pm)-(d.ex+bo)));
    const s=sug(D);const e=P.empaque(s,ctx.emp.get(k),d.unidad);
    // MF
    const smf=P.serie(ctx.ventas.MF.get(k),null);
    const apm=P.antPron(smf.s,ref,N,m=>ctx.agot.has('MF|'+k+'|'+m),ctx.ultimo.MF,ctx.primero&&ctx.primero.MF);
    const Dm=apm.I!=null?apm.I:(apm.A!=null&&apm.P!=null?(apm.A+apm.P)/2:(apm.A??apm.P??0));
    const inv=ctx.invMF.has(k)?ctx.invMF.get(k):null;
    const mm=inv!=null&&Dm>0?inv/Dm:null;const flag=mm!=null&&mm>P.R.MF_MESES_REV;
    const r1=x=>x==null?null:Math.round(x*10)/10;
    const notas=[...(sv.ed.length?['ED restada '+sv.ed.join(',')]:[]),...ap.notas,...(ap.I!=null?['venta intermitente: demanda = promedio 12 meses']:ap.P==null&&ap.A!=null?['sin pronóstico válido: demanda = antecedente']:ap.A==null&&ap.P!=null?['sin antecedente válido: demanda = pronóstico']:[]),...(e.nota?[e.nota]:[])];
    const eX=(X,pr)=>X==null?null:P.estatus(d.rot,d.ex,X,lt,pm,pr);
    out.push({c:k,n:d.nombre,m:d.marca||'',comp:d.comprador,rot:d.rot,g:d.grupo,u:d.unidad,lt,pm,N,ex:d.ex,bo,boT:d.bo,eA:eX(ap.A,pv.eA),eP:eX(ap.P,pv.eP),
      A:r1(ap.A),P:r1(ap.P),D:r1(D),bD:ap.I!=null?'INTERMITENTE 12 MESES':ap.A!=null&&ap.P!=null?'PROMEDIO ANT+PRON':ap.A!=null?'SOLO ANTECEDENTE':ap.P!=null?'SOLO PRONÓSTICO':'SIN VENTAS',mi:D>0?Math.round(d.ex/D*100)/100:null,miB:D>0?Math.round((d.ex+bo)/D*100)/100:null,est,estX:d.estX,
      sA:sug(ap.A),sP:sug(ap.P),s,sE:e.ok?e.q:null,eT:e.ok?e.tipo:null,
      mfA:r1(apm.A),mfP:r1(apm.P),mfD:r1(Dm),mfI:inv,mfM:mm==null?null:Math.round(mm*100)/100,
      mfF:flag?1:0,mfX:flag?Math.floor(inv-Dm*pm):null,nt:notas.join(' · ')});
  });
  return out;
};
const EN_RIESGO=['RIESGO','FALTANTE'];
// ── Explicación mes a mes de una clave (tablero): qué meses entraron al antecedente y al pronóstico y por qué se descartaron otros.
// Usa exactamente las mismas reglas que P.antPron (hay prueba que lo verifica).
P.explicar=(meses,ed,ref,N,agot,ultimo,primero)=>{
  const sv=P.serie(meses,ed);const s=sv.s;const fin=ultimo&&ultimo<P.ym(ref,-1)?ultimo:P.ym(ref,-1);
  const conVenta=Object.keys(s).filter(m=>s[m]>0).sort();const ini=conVenta[0]||null;const hi=ultimo||fin;
  const por=m=>ini==null||m<ini?'antes de la primera venta':agot(m)?'agotado (más de '+(P.R.DIAS_AGOTADO-1)+' días sin existencia)':P.sospechoso(s,m,primero,hi)?'cero sospechoso de agotado':null;
  const out=new Map();
  const reg=(m,win)=>{
    const o=out.get(m)||{m,venta:meses&&meses[m]!=null?meses[m]:null,limpia:s[m]!=null?s[m]:null,ed:sv.ed.includes(m),usa:[]};
    const x=por(m);if(x){o.excluido=x;out.set(m,o);return false;}
    const r=P.valor(s,m);if(r.at)o.atipico=r.at;if(r.v==null){o.excluido='atípico descartado (>'+P.R.ATIPICO+'× mediana)';out.set(m,o);return false;}
    o.valor=r.v;if(win)o.usa.push(win);out.set(m,o);return true;};
  let n=0;for(let i=0;i<P.R.VENTANA_ATRAS&&n<N;i++){const m=P.ym(fin,-i);if(reg(m,'ANT'))n++;}
  for(let i=0;i<N;i++){const m=P.ym(ref,i-12);if(ultimo&&m>ultimo)continue;reg(m,'PRON');}
  for(let i=1;i<=24;i++){const m=P.ym(ref,-i);if(!out.has(m)&&(!ultimo||m<=ultimo))reg(m,null);}
  return[...out.values()].sort((a,b)=>a.m<b.m?-1:1);
};
// Fracción del BO que llega al CEDI por clave = pendiente de OCs al CEDI ÷ pendiente total (CEDI + entrega directa).
// Se usa proporción porque la cantidad de la OC puede venir en unidad de compra y el BO en unidad de inventario.
P.fracBoCEDI=(cur,enc,sinED)=>{
  const g=new Map();
  (cur?[...cur.values()]:[]).forEach(l=>{const p=Number(l.cantidad_pend)||0;if(p<=0)return;const e=enc&&enc.get(Number(l.folio));const ed=P.esED(e&&e.destino,l.comprador||(e&&e.comprador),sinED);
    const o=g.get(l.clave)||{t:0,ed:0};o.t+=p;if(ed)o.ed+=p;g.set(l.clave,o);});
  const m=new Map();g.forEach((o,k)=>{if(o.ed>0)m.set(k,(o.t-o.ed)/o.t);});return m;
};
// Conteo de claves por estatus para la tendencia del tablero (una fila por comprador, base y estatus)
P.histEstatus=(res,fecha)=>{
  const g={};const add=(c,b,e)=>{if(!e)return;const k=c+'|'+b+'|'+e;g[k]=(g[k]||0)+1;};
  res.forEach(r=>{const c=r.comp||'(SIN COMPRADOR)';add(c,'PORTAL',r.est);add(c,'ANT',r.eA||'SIN CONSUMO');add(c,'PRON',r.eP||'SIN CONSUMO');add(c,'EXCEL',r.estX||'(VACÍO)');});
  return Object.entries(g).map(([k,n])=>{const[c,b,e]=k.split('|');return{fecha,comprador:c,base:b,estatus:e,claves:n};});
};
// Lo que sugería el portal (último cálculo publicado) cuando aparece una OC nueva
// (siempre trae todas las columnas: en un upsert por lotes una columna faltante se guardaría como vacía)
P.sugOC=(r,fecha)=>({sug_fecha:r?fecha:null,sug_portal:r?r.s:null,sug_empaque:r?r.sE:null,sug_ant:r?r.sA:null,sug_pron:r?r.sP:null,sug_estatus:r?r.est:null,sug_demanda:r?r.D:null});
// Pedido vs sugerencia del portal. ped y sug en la unidad de inventario. Tolerancia ±40% (la misma del tablero actual de Pedidos propuestos)
P.R.TOL_PEDIDO=40;
P.precisionOC=(ped,sug)=>{
  if(sug==null||ped==null)return{res:'NO EVALUABLE',desv:null};
  if(sug<=0)return{res:ped>0?'PEDIDO SIN SUGERENCIA':'ACERTADA',desv:null};
  const d=Math.round((ped-sug)/sug*100);
  return{res:Math.abs(d)<=P.R.TOL_PEDIDO?'ACERTADA':'DESVIADA',desv:d};
};
// Compara contra el día anterior: claves que ENTRAN a riesgo (alarma) y las que salen (cierre)
P.alarmas=(prev,prevEst,res,fecha,restr)=>{
  const nuevas=[],cierres=[],set=[];
  res.forEach(r=>{
    const enR=EN_RIESGO.includes(r.est);if(enR)set.push(r.c);
    if(enR&&!prev.has(r.c)){
      if(P.R.EXCLUIR_ROT.includes(r.rot)||!r.D)return;
      const directo=r.est==='FALTANTE';
      nuevas.push({clave:r.c,comprador:r.comp,rotacion:r.rot,fecha_alarma:fecha,tipo:directo?'FALTANTE_DIRECTO':'RIESGO',
        n_meses:r.N,t_entrega:r.lt,objetivo_meses:r.pm,antecedente:r.A,pronostico:r.P,demanda:r.D,existencias:r.ex,bo:r.bo,meses_inv:r.mi,
        sugerida:r.s,sugerida_empaque:r.sE,resultado_target:restr.has(r.c+'|'+r.comp)?'EXCLUIDA':(r.mfF&&r.s>0&&r.mfX>=r.s?'EN REVISION GERENCIA':null),
        nota:[directo?'Posible venta atípica (pasó directo a FALTANTE)':'',restr.has(r.c+'|'+r.comp)?'Clave restringida':'',r.mfF&&r.s>0&&r.mfX>=r.s?'Excedente MF ('+r.mfX+') cubre la sugerencia: revisar redistribución con gerencia':r.mfF?'MF con más de '+P.R.MF_MESES_REV+' meses, pero su excedente no cubre la sugerencia':''].filter(Boolean).join(' · ')||null});
    }
  });
  const cur=new Set(set);prev.forEach(k=>{if(!cur.has(k))cierres.push(k);});
  return{nuevas,cierres,set};
};
// ══ OCs ══
// Líneas pendientes de "Productos Incluidos en" → Map 'folio|clave' → línea
P.ocLineas=wb=>{
  const ws=P.hoja(wb,['Productos Incluidos en']);if(!ws)return null;
  const aoa=P.aoa(ws);const h=P.header(aoa,['FOLIO OC','FECHA OC','CLAVE','CANTIDAD PED','CANTIDAD PEND','FECHA ENTREGA']);
  if(!h)throw new Error('La hoja Productos Incluidos en no trae FOLIO OC, FECHA OC, CLAVE, CANTIDAD PED, CANTIDAD PEND y FECHA ENTREGA.');
  const iP=h.cols.indexOf('PROVEEDOR'),iC=h.cols.indexOf('COMPRADOR'),iE=h.cols.indexOf('ESTADO OC');
  const m=new Map();
  for(let i=h.fila+1;i<aoa.length;i++){
    const r=aoa[i];const f=P.clave(r[h.idx['FOLIO OC']]),k=P.clave(r[h.idx.CLAVE]);if(f==null||k==null)continue;
    const key=f+'|'+k;const ped=P.num(r[h.idx['CANTIDAD PED']])||0,pend=P.num(r[h.idx['CANTIDAD PEND']])||0;
    const prev=m.get(key);
    if(prev){prev.cantidad_ped+=ped;prev.cantidad_pend+=pend;continue;}   // misma clave repetida en la OC: se suma
    m.set(key,{folio:f,clave:k,fecha_oc:P.fecha(r[h.idx['FECHA OC']]),proveedor:iP>=0?String(r[iP]).trim():null,comprador:iC>=0?P.norm(r[iC]):null,
      cantidad_ped:ped,cantidad_pend:pend,fecha_entrega:P.fecha(r[h.idx['FECHA ENTREGA']]),estado_oc:iE>=0?P.norm(r[iE]):null});
  }
  return m;
};
// ── Entrega directa: la sucursal de destino aparece en Observaciones de RP01-23 ──
// Cuenta si la sucursal está al inicio (opcionalmente tras "MAS FERRE", "SANVER", "CIUDAD/CD"),
// o después de "ENTREGAR EN", "SUC./SUCURSAL" o "PARA". Así se evitan falsos como
// "CEMENTO BLANCO MORA" (marca) o "RECOLECCION TAMPICO".
P.R.SANVER_SUC_ES_ED=false;  // "SANVER ALAMO", "SANVER LA FUENTE": entran al CEDI (confirmado por Compras, 14-oct-2026)
P.SUCURSALES=['AHUATENO','ALAMO','ARISTA','ATLAPEXCO','BODEGA','CENTRO','CHICONTEPEC','COATZINTLA','DEL CARMEN','EL CARMEN','EBANO','LA FUENTE','LA RIVERA','MANTE','MARTINEZ','MERCADO','MORA','NARANJOS','PINTURAS','PUEBLO VIEJO','TAMIAHUA','TAMPICO','TEMPOAL','TIHUATLAN','TUXPAN','VALLES','VALLEZ','ZACATIPAN'];
P.destinoED=obs=>{
  const o=P.norm(obs).replace(/^[^A-Z0-9]+/,'');if(!o)return null;
  const ALT='(?:CIUDAD |CD\\.? )?('+P.SUCURSALES.join('|')+')(?![A-Z])',SUC='(?:SUC(?:URSAL)?\\.? )?';
  const pre='(?:MAS FERRE(?:TERIAS)?\\.? (?:SANVER )?|MF )?'+(P.R.SANVER_SUC_ES_ED?'(?:SANVER (?!FORTE))?':'');
  const rx=[new RegExp('^'+pre+SUC+ALT),new RegExp('ENTREGA(?:R)? (?:EN )?'+SUC+ALT),new RegExp('(?:^|[^A-Z])SUC(?:URSAL)?\\.? '+ALT),new RegExp('(?:^|[^A-Z])PARA (?:LA )?'+SUC+ALT)];
  for(const r of rx){const m=o.match(r);if(m)return m[1]==='VALLEZ'?'VALLES':m[1]==='EL CARMEN'?'DEL CARMEN':m[1];}
  return null;
};
// Compradores que reciben en sucursal y después mueven la mercancía al CEDI: sus OCs NO son entrega directa
// aunque Observaciones traiga una sucursal. La lista se edita en Cálculo diario (tabla plan_config); esta es la de arranque.
P.R.SIN_ED_DEF=['IVAN','MYRIAM VIDA CHRISTELL','KARLA ANAID','JISHELL MARINTHYA'];
P.esED=(destino,comprador,sinED)=>!!destino&&!(sinED||new Set(P.R.SIN_ED_DEF)).has(P.norm(comprador||''));
// Encabezados de OC (RP01-23): Map folio → {estatus, estado, obs}
P.ocEncabezados=wb=>{
  const ws=P.hoja(wb,['RP01-23 Ordenes de Compra']);if(!ws)return new Map();
  const aoa=P.aoa(ws);const h=P.header(aoa,['FOLIO','ESTATUS']);if(!h)return new Map();
  const iE=h.cols.indexOf('ESTADO OC'),iO=h.cols.indexOf('OBSERVACIONES'),iC=h.cols.indexOf('COMPRADOR');const m=new Map();
  for(let i=h.fila+1;i<aoa.length;i++){const r=aoa[i];const f=P.clave(r[h.idx.FOLIO]);if(f==null)continue;const obs=iO>=0?String(r[iO]||'').trim():'';const d=P.destinoED(obs);m.set(f,{estatus:P.norm(r[h.idx.ESTATUS]),estado:iE>=0?P.norm(r[iE]):'',obs,destino:d,comprador:iC>=0?P.norm(r[iC]):''});}
  return m;
};
// Compara la foto anterior (líneas abiertas en la base) con la de hoy
P.diffOC=(prev,cur,enc,fecha)=>{
  const upserts=[],recep=[],r4=x=>Math.round(x*10000)/10000;
  cur.forEach((l,key)=>{
    const p=prev.get(key);
    if(!p){upserts.push({...l,primera_foto:fecha,pend_inicial:l.cantidad_pend,cancelada:false,cerrada_en:null});return;}
    const baja=Number(p.cantidad_pend)-l.cantidad_pend;
    if(baja>1e-9)recep.push({folio:l.folio,clave:l.clave,fecha,cantidad:r4(baja),fecha_entrega:l.fecha_entrega});
    upserts.push({...l,primera_foto:p.primera_foto,pend_inicial:p.pend_inicial,cancelada:false,cerrada_en:null});
  });
  let canc=0,comp=0;
  prev.forEach((p,key)=>{
    if(cur.has(key))return;
    const e=enc.get(Number(p.folio));const cancel=!!(e&&e.estatus==='CANCELADO');
    if(cancel)canc++;else{comp++;if(Number(p.cantidad_pend)>0)recep.push({folio:Number(p.folio),clave:Number(p.clave),fecha,cantidad:r4(Number(p.cantidad_pend)),fecha_entrega:p.fecha_entrega});}
    upserts.push({folio:Number(p.folio),clave:Number(p.clave),fecha_oc:p.fecha_oc,proveedor:p.proveedor,comprador:p.comprador,cantidad_ped:Number(p.cantidad_ped),
      cantidad_pend:cancel?Number(p.cantidad_pend):0,fecha_entrega:p.fecha_entrega,estado_oc:cancel?'CANCELADO':'RECIBIDO TOTALMENTE',entrega_directa:!!p.entrega_directa,destino:p.destino||null,observaciones:p.observaciones||null,
      primera_foto:p.primera_foto,pend_inicial:p.pend_inicial,cancelada:cancel,cerrada_en:fecha});
  });
  return{upserts,recep,cerradas:comp,canceladas:canc};
};
P.d10=v=>v==null||v===''?null:String(v).slice(0,10);   // timestamptz de Supabase → 'AAAA-MM-DD'
P.dias=(a,b)=>Math.round((new Date(P.d10(b)+'T12:00:00')-new Date(P.d10(a)+'T12:00:00'))/864e5);
// Target de una alarma. ocs: líneas de esa clave {folio,fecha_oc,cantidad_pend,cerrada_en,cancelada}
P.R.T_ANTES=15;P.R.T_ANTES_PCT=0.10;P.R.T_DESPUES=5;P.R.T_VIEJA=60;P.R.DIAS_MES=30;P.R.MF_MESES_REV=8;
// Margen para considerar una OC anticipada: 15 días o el 10% del tiempo de entrega, lo que sea mayor
P.margen=lt=>Math.max(P.R.T_ANTES,Math.round((Number(lt)||0)*P.R.DIAS_MES*P.R.T_ANTES_PCT));
P.target=(al,ocs,hoy)=>{
  const fa=P.d10(al.fecha_alarma);const vs=ocs.filter(o=>o.fecha_oc&&!o.cancelada&&!o.entrega_directa).map(o=>({...o,fecha_oc:P.d10(o.fecha_oc),d:P.dias(fa,o.fecha_oc)})).sort((a,b)=>a.d-b.d);
  const mg=P.margen(al.t_entrega),vieja0=Math.max(P.R.T_VIEJA,Math.round((Number(al.t_entrega)||0)*P.R.DIAS_MES));
  const enRango=vs.find(o=>o.d>=-mg&&o.d<=P.R.T_DESPUES);
  const res=(r,o)=>({resultado_target:r,oc_folio:o?Number(o.folio):null,fecha_oc:o?o.fecha_oc:null,dias_dif:o?o.d:null});
  if(enRango)return res(enRango.d<0?'ANTICIPADA':'EN TIEMPO',enRango);
  const vieja=vs.filter(o=>o.d<-mg&&o.d>=-vieja0&&(!o.cerrada_en||o.cerrada_en>=fa)).pop();
  if(vieja)return res('ANTICIPACION EXCESIVA',vieja);
  const tarde=vs.find(o=>o.d>P.R.T_DESPUES);
  if(tarde)return res('TARDE',tarde);
  if(al.fecha_fin&&P.dias(fa,al.fecha_fin)<=P.R.T_DESPUES)return res('SIN NECESIDAD',null);
  if(P.dias(fa,hoy)>P.R.T_DESPUES)return res('SIN OC',null);   // provisional: cambia a TARDE si llega la OC
  return null;   // todavía dentro del plazo
};
// OTIF por línea cerrada. compromisos: Map folio → fecha
P.otifLinea=(l,recs,compromisos,lt)=>{
  if(l.entrega_directa)return null;                                 // entrega directa a sucursal: no es OTIF del CEDI
  if(l.cancelada&&!(recs&&recs.length))return null;              // cancelada sin recibir: no cuenta
  if(Number(l.pend_inicial)<Number(l.cantidad_ped)-1e-9)return null; // ya traía recepciones antes de la primera foto
  const rec=(recs||[]).slice().sort((a,b)=>a.fecha<b.fecha?-1:1);if(!rec.length)return null;
  // Fecha compromiso: la capturada por el comprador; si no hay, la de POR DEFECTO = fecha de la OC + meses de entrega
  // de la clave × 30 días. Solo si la clave no tiene tiempo de entrega se usa la fecha de entrega de la OC (si es válida).
  let ref=null,base='SIN FECHA';
  if(compromisos.has(Number(l.folio))){ref=compromisos.get(Number(l.folio));base='COMPROMISO CAPTURADO';}
  else if(l.fecha_oc&&Number(lt)>0){const d=new Date(l.fecha_oc+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+Math.round(Number(lt)*P.R.DIAS_MES));ref=d.toISOString().slice(0,10);base='COMPROMISO POR DEFECTO';}
  else if(l.fecha_entrega&&(!l.fecha_oc||l.fecha_entrega>l.fecha_oc)){ref=l.fecha_entrega;base='ENTREGA EN OC (sin tiempo de entrega)';}
  const tot=rec.reduce((t,r)=>t+Number(r.cantidad),0);const ult=rec[rec.length-1].fecha;
  return{folio:Number(l.folio),clave:Number(l.clave),proveedor:l.proveedor,comprador:l.comprador,mes:ult.slice(0,7),ref,base,
    on_time:ref?ult<=ref:null,in_full:!l.cancelada&&tot>=Number(l.cantidad_ped)-1e-9,recibido:tot,pedido:Number(l.cantidad_ped)};
};
P.otifResumen=(lineas,clave)=>{
  // On Time solo sobre líneas con fecha de referencia válida; In Full sobre todas. OTIF = promedio de los dos.
  const g={};lineas.forEach(x=>{const k=x[clave]||'(sin dato)';const o=g[k]||(g[k]={n:0,nOT:0,ot:0,inf:0,comp:0});o.n++;if(x.on_time!=null){o.nOT++;if(x.on_time)o.ot++;}if(x.in_full)o.inf++;if(x.base==='COMPROMISO CAPTURADO')o.comp++;});
  return Object.entries(g).map(([k,o])=>{const ot=o.nOT?o.ot/o.nOT:null,inf=o.inf/o.n;return{nombre:k,lineas:o.n,con_fecha:o.nOT,on_time:ot,in_full:inf,otif:ot==null?null:(ot+inf)/2,con_compromiso:o.comp/o.n};}).sort((a,b)=>b.lineas-a.lineas);
};
// Temporada: variación fuerte entre antecedente y pronóstico (pronóstico = mismos meses del año pasado que vienen).
// SUBE: el pronóstico es 50% (ajustable en pantalla) o más arriba del antecedente (y al menos 1 unidad/mes más) · BAJA: el mismo % o más abajo · ESTABLE: el resto.
P.R.TEMP_VAR=0.50;P.R.TEMP_MIN=1;
P.temporada=(A,Pr)=>{
  if(A==null||Pr==null)return{t:'SIN DATO',v:null};
  if(!(A>0))return Pr>=P.R.TEMP_MIN?{t:'SUBE',v:null}:{t:'ESTABLE',v:null};
  const v=(Pr-A)/A;
  if(v>=P.R.TEMP_VAR&&Pr-A>=P.R.TEMP_MIN)return{t:'SUBE',v};
  if(v<=-P.R.TEMP_VAR&&A-Pr>=P.R.TEMP_MIN)return{t:'BAJA',v};
  return{t:'ESTABLE',v};
};
P.TEMP_L={SUBE:'↑ Sube fuerte (temporada)',BAJA:'↓ Baja fuerte',ESTABLE:'Sin cambio fuerte','SIN DATO':'Sin antecedente o pronóstico'};
/*PARSE-END*/
if(typeof window!=='undefined')window.P=P;   // para páginas que preguntan por window.P (Estatus Back Orders)
