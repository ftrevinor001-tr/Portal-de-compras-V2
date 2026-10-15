// Portal de Compras V2 · Mi Resumen con el cálculo del portal (Planeación e indicadores)
// Reemplaza en Mi Resumen las tarjetas de Riesgo, Sobre inventario, Pedidos propuestos y Target
// (que antes salían del Excel) y agrega OTIF de proveedor. Usa P de plan_logica.js.
(function(){
  const d10=v=>v==null||v===''?null:String(v).slice(0,10);
  async function todo(q){let out=[];for(let i=0;i<200;i++){const{data,error}=await q().range(i*1000,i*1000+999);if(error)throw error;out=out.concat(data||[]);if(!data||data.length<1000)break;}return out;}
  const r1=x=>x==null?'':(Math.round(x*10)/10);
  const CUMPLE=['EN TIEMPO','ANTICIPADA'],NO_CUENTA=['EXCLUIDA','EN REVISION GERENCIA','SIN NECESIDAD'];

  async function cargar(sb){
    const{data:pub,error}=await sb.from('plan_publicacion').select('fecha,ruta').eq('id','estado').maybeSingle();
    if(error||!pub)throw new Error(error?error.message:'sin cálculo publicado');
    const{data:blob,error:e2}=await sb.storage.from('respaldos').download(pub.ruta);if(e2)throw e2;
    const txt=pub.ruta.endsWith('.gz')?await new Response(blob.stream().pipeThrough(new DecompressionStream('gzip'))).text():await blob.text();
    const est=JSON.parse(txt);const mes=est.fecha.slice(0,7);const ini=mes+'-01';
    const [al,cer,cmp]=await Promise.all([
      todo(()=>sb.from('plan_alarmas').select('clave,comprador,fecha_alarma,tipo,resultado_target,oc_folio,dias_dif').gte('fecha_alarma',ini).order('fecha_alarma').order('clave')),
      todo(()=>sb.from('plan_oc_lineas').select('folio,clave,fecha_oc,proveedor,comprador,cantidad_ped,cantidad_pend,fecha_entrega,cancelada,cerrada_en,pend_inicial,entrega_directa').gte('cerrada_en',ini).order('folio').order('clave')),
      todo(()=>sb.from('plan_oc_compromiso').select('folio,fecha').order('folio'))]);
    const folios=[...new Set(cer.map(l=>Number(l.folio)))];const recs=new Map();
    for(let i=0;i<folios.length;i+=200){(await todo(()=>sb.from('plan_recepciones').select('folio,clave,fecha,cantidad').in('folio',folios.slice(i,i+200)).order('folio').order('clave').order('fecha'))).forEach(r=>{const k=Number(r.folio)+'|'+Number(r.clave);if(!recs.has(k))recs.set(k,[]);recs.get(k).push({...r,fecha:d10(r.fecha)});});}
    const comp=new Map(cmp.filter(c=>c.fecha).map(c=>[Number(c.folio),d10(c.fecha)]));
    const lt=new Map((est.claves||[]).map(r=>[Number(r.c),r.lt]));
    const otif=cer.map(l=>P.otifLinea({...l,fecha_oc:d10(l.fecha_oc),fecha_entrega:d10(l.fecha_entrega)},recs.get(Number(l.folio)+'|'+Number(l.clave)),comp,lt.get(Number(l.clave)))).filter(x=>x&&x.mes===mes);
    return{fecha:est.fecha,mes,claves:est.claves||[],ocs:est.ocs||{},alarmas:al.map(a=>({...a,fecha_alarma:d10(a.fecha_alarma)})),otif};
  }

  function aplicar(alarmas,v){
    const TIPOS_V2=['RIESGO','SOBREINVENTARIO','PEDIDO_PROPUESTO','TARGET','OTIF'];
    const comps=[...new Set([...alarmas.map(a=>a.comprador),...v.claves.map(r=>r.comp).filter(Boolean)])];if(!comps.includes('GENERAL'))comps.unshift('GENERAL');
    const out=alarmas.filter(a=>!TIPOS_V2.includes(a.tipo));
    const nota='Cálculo del portal al '+v.fecha.split('-').reverse().join('/')+' (Planeación e indicadores).';
    const push=(comp,tipo,items,total,descripcion,extra)=>out.push({comprador:comp,dashboard:'planeacion',tipo,mes:v.mes,descripcion,severidad:null,datos:Object.assign({total,items:items.slice(0,300)},extra||{})});
    comps.forEach(comp=>{
      const mine=x=>comp==='GENERAL'||x===comp;
      const cl=v.claves.filter(r=>mine(r.comp));
      const rsg=cl.filter(r=>r.est==='RIESGO').sort((a,b)=>(a.mi??0)-(b.mi??0)).map(r=>({Clave:r.c,Nombre:r.n,Marca:r.m||'','Meses Inv':r1(r.mi),'T.Entrega':r.lt,'Demanda':r1(r.D),'Propuesta':r.sE!=null?r.sE:r.s}));
      push(comp,'RIESGO',rsg,rsg.length,nota+' Meses de inventario por debajo del tiempo de entrega × 1.4 con la demanda del portal (promedio de antecedente y pronóstico).');
      const sob=cl.filter(r=>r.est==='SOBRE INVENTARIO').sort((a,b)=>(b.mi??0)-(a.mi??0)).map(r=>({Clave:r.c,Nombre:r.n,Marca:r.m||'','Meses Inv':r1(r.mi),'Obj':r.pm,'Demanda':r1(r.D)}));
      push(comp,'SOBREINVENTARIO',sob,sob.length,nota+' Más meses de inventario que el objetivo × 1.4.');
      const ped=cl.filter(r=>(r.est==='RIESGO'||r.est==='FALTANTE')&&(r.s||0)>0).sort((a,b)=>(a.mi??-1)-(b.mi??-1)).map(r=>({Clave:r.c,Descripción:r.n,Estatus:r.est,'Propuesta':r.s,'Con empaque':r.sE!=null?r.sE:'','Ant.':r.sA!=null?r.sA:'','Pron.':r.sP!=null?r.sP:'','OCs abiertas':(v.ocs[r.c]||[]).length}));
      push(comp,'PEDIDO_PROPUESTO',ped,ped.length,nota+' Faltante y riesgo con propuesta = demanda × (tiempo de entrega + objetivo) − (existencia + BO del CEDI, sin entregas directas).');
      const al=v.alarmas.filter(a=>mine(a.comprador));let n=0,ok=0,pend=0;
      al.forEach(a=>{if(!a.resultado_target){pend++;return;}if(NO_CUENTA.includes(a.resultado_target))return;n++;if(CUMPLE.includes(a.resultado_target))ok++;});
      const ord={'TARDE':0,'SIN OC':1,'ANTICIPACION EXCESIVA':2};
      const itT=al.slice().sort((a,b)=>(ord[a.resultado_target]??5)-(ord[b.resultado_target]??5)).map(a=>({Clave:a.clave,'Fecha alarma':a.fecha_alarma.split('-').reverse().join('/'),Resultado:a.resultado_target||'EN PLAZO',OC:a.oc_folio||'','Días':a.dias_dif==null?'':a.dias_dif,...(comp==='GENERAL'?{Comprador:a.comprador}:{})}));
      push(comp,'TARGET',itT,n,nota+' Cumple si la OC se coloca en tiempo o anticipada respecto al día en que la clave entró a riesgo.',{ocsOk:ok,pct:n?Math.round(ok/n*1000)/10:null,texto:ok+' de '+n+' alarmas evaluadas cumplen Target (<b>'+(n?(Math.round(ok/n*1000)/10)+'%':'—')+'</b>) · '+pend+' aún en plazo.'});
      const ol=v.otif.filter(l=>mine(l.comprador));
      const g=ol.length?P.otifResumen(ol.map(x=>({...x,g:'G'})),'g')[0]:null;
      const pc=x=>x==null?'—':(Math.round(x*1000)/10)+'%';
      const itO=P.otifResumen(ol,'proveedor').sort((a,b)=>(a.otif??2)-(b.otif??2)).map(o=>({Proveedor:o.nombre,'Líneas':o.lineas,'On Time':pc(o.on_time),'In Full':pc(o.in_full),OTIF:pc(o.otif)}));
      push(comp,'OTIF',itO,ol.length,nota+' Líneas de OC cerradas en el mes; On Time contra la fecha compromiso (capturada o fecha OC + tiempo de entrega).',{pct:g&&g.otif!=null?Math.round(g.otif*1000)/10:null,texto:ol.length?'OTIF <b>'+pc(g.otif)+'</b> · On Time '+pc(g.on_time)+' · In Full '+pc(g.in_full)+' · '+ol.length+' líneas cerradas.':'Aún no hay líneas de OC cerradas este mes.'});
    });
    return out;
  }
  window.RESUMEN_V2={cargar,aplicar};
})();
