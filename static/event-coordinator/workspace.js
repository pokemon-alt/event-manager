
(() => {
  const root=document.getElementById('event-coordinator');
  const privacyKey='event-coordinator.privacy-mode';
  let privacyMode=false;
  try{privacyMode=localStorage.getItem(privacyKey)==='on';}catch{}
  function ecPrivacyEnabled(){return privacyMode;}
  function ecPrivateText(value){return privacyMode&&value?'••••':String(value??'');}
  function syncPrivacyControls(){
    root.dataset.privacyMode=privacyMode?'on':'off';
    for(const button of root.querySelectorAll('[data-ec-privacy-toggle]')){
      button.setAttribute('aria-pressed',String(privacyMode));
      button.querySelector('span').textContent=privacyMode?'On':'Off';
    }
    for(const input of root.querySelectorAll('input[data-ec-private]'))input.type=privacyMode?'password':'text';
  }
  function setPrivacyMode(enabled){
    if(privacyMode===enabled)return;
    privacyMode=enabled;
    syncPrivacyControls();
    root.dispatchEvent(new CustomEvent('ec:privacy-change',{detail:enabled}));
  }
  for(const button of root.querySelectorAll('[data-ec-privacy-toggle]'))button.addEventListener('click',()=>{
    setPrivacyMode(!privacyMode);
    try{localStorage.setItem(privacyKey,privacyMode?'on':'off');}catch{}
  });
  window.addEventListener('storage',event=>{
    if(event.key===privacyKey||event.key===null)setPrivacyMode(event.newValue==='on');
  });
  syncPrivacyControls();
  const mappedImportAdapters=Object.create(null);
  let privateExportFile=null;
    // A single-sheet OOXML workbook. Shared strings and Excel's built-in Text style
  // preserve identifiers, including when an exported cell is edited in Excel.
  function makeTextWorkbook(matrix) {
    const sheetNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const docRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const packageRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const indices = new Map();
    const strings = [];
    const columns = [];
    const widths = [];
    let cellCount = 0;

    function columnName(index) {
      let name = '';
      for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
        name = String.fromCharCode(65 + (n - 1) % 26) + name;
      }
      return name;
    }

    function stringIndex(value) {
      if (typeof value !== 'string') throw new Error('A workbook field must be text.');
      if (value.length > 32767) throw new Error('A field exceeds Excel’s 32,767-character limit.');
      if (indices.has(value)) return indices.get(value);
      // Escape literal OOXML escape sequences first. XML carriage returns need
      // a character reference to survive the XML parser's newline normalization.
      const xml = value.replace(/_(?=x[0-9a-f]{4}_)/gi, '_x005F_')
        .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]/gu,
          c => '_x' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#13;');
      const index = strings.length;
      indices.set(value, index);
      strings.push('<si><t xml:space="preserve">' + xml + '</t></si>');
      return index;
    }

    if (matrix.length > 1048576) throw new Error('This list exceeds Excel’s row limit.');
    const rows = matrix.map((row, index) => {
      if (row.length > 16384) throw new Error('This list exceeds Excel’s column limit.');
      const rowNumber = index + 1;
      const cells = row.map((value, col) => {
        const stringId = stringIndex(value);
        if (!columns[col]) columns[col] = columnName(col);
        widths[col] = Math.max(widths[col] || 12, Math.min(value.length + 3, 52));
        cellCount++;
        return '<c r="' + columns[col] + rowNumber + '" s="1" t="s"><v>' + stringId + '</v></c>';
      });
      return '<row r="' + rowNumber + '">' + cells.join('') + '</row>';
    });
    const cols = widths.map((width, col) =>
      '<col min="' + (col + 1) + '" max="' + (col + 1) + '" width="' + width + '" customWidth="1" style="1"/>');
    const sheet = '<worksheet xmlns="' + sheetNS + '"><sheetViews><sheetView workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      (cols.length ? '<cols>' + cols.join('') + '</cols>' : '') + '<sheetData>' + rows.join('') + '</sheetData></worksheet>';
    const styles = '<styleSheet xmlns="' + sheetNS + '">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="49" applyNumberFormat="1" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    const parts = [
      ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="' + contentType + 'sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + contentType + 'worksheet+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="' + contentType + 'sharedStrings+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="' + contentType + 'styles+xml"/></Types>'],
      ['_rels/.rels', '<Relationships xmlns="' + packageRel + '"><Relationship Id="rId1" Type="' +
        docRel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml', '<workbook xmlns="' + sheetNS + '" xmlns:r="' + docRel + '">' +
        '<bookViews><workbookView/></bookViews><sheets><sheet name="Profiles" sheetId="1" r:id="rId1"/></sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="' + packageRel + '">' +
        '<Relationship Id="rId1" Type="' + docRel + '/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="' + docRel + '/sharedStrings" Target="sharedStrings.xml"/>' +
        '<Relationship Id="rId3" Type="' + docRel + '/styles" Target="styles.xml"/></Relationships>'],
      ['xl/worksheets/sheet1.xml', sheet],
      ['xl/styles.xml', styles],
      ['xl/sharedStrings.xml', '<sst xmlns="' + sheetNS + '" count="' + cellCount + '" uniqueCount="' +
        strings.length + '">' + strings.join('') + '</sst>']
    ];
    return zipWorkbookParts(parts.map(([name, xml]) => [name, declaration + xml]));
  }

  // ZIP "store" method. All filenames are fixed ASCII package paths.
  function zipWorkbookParts(parts) {
    const encoder = new TextEncoder();
    const crcTable = Uint32Array.from({length: 256}, (_, value) => {
      for (let bit = 0; bit < 8; bit++) value = (value >>> 1) ^ ((value & 1) ? 0xedb88320 : 0);
      return value;
    });
    const files = parts.map(([path, xml]) => {
      const bytes = encoder.encode(xml);
      let crc = 0xffffffff;
      for (const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 255];
      return {name: encoder.encode(path), bytes, crc: (crc ^ 0xffffffff) >>> 0};
    });
    const directoryStart = files.reduce((size, file) => size + 30 + file.name.length + file.bytes.length, 0);
    const directorySize = files.reduce((size, file) => size + 46 + file.name.length, 0);
    const archiveSize = directoryStart + directorySize + 22;
    if (archiveSize > 0xffffffff) throw new Error('This workbook exceeds the supported ZIP size.');
    const archive = new Uint8Array(archiveSize);
    const view = new DataView(archive.buffer);
    let local = 0;
    let directory = directoryStart;
    for (const {name, bytes, crc} of files) {
      view.setUint32(local, 0x04034b50, true);
      view.setUint16(local + 4, 20, true);
      view.setUint16(local + 12, 33, true); // DOS date: 1980-01-01.
      view.setUint32(local + 14, crc, true);
      view.setUint32(local + 18, bytes.length, true);
      view.setUint32(local + 22, bytes.length, true);
      view.setUint16(local + 26, name.length, true);
      archive.set(name, local + 30);
      archive.set(bytes, local + 30 + name.length);

      view.setUint32(directory, 0x02014b50, true);
      view.setUint16(directory + 4, 20, true);
      // The common fields from 'version needed' through 'extra field length'
      // have the same layout in both the local header and the central directory.
      archive.set(archive.subarray(local + 4, local + 30), directory + 6);
      view.setUint32(directory + 42, local, true);
      archive.set(name, directory + 46);
      local += 30 + name.length + bytes.length;
      directory += 46 + name.length;
    }
    view.setUint32(directory, 0x06054b50, true);
    view.setUint16(directory + 8, files.length, true);
    view.setUint16(directory + 10, files.length, true);
    view.setUint32(directory + 12, directorySize, true);
    view.setUint32(directory + 16, directoryStart, true);
    return archive;
  }

  const get=id=>root.querySelector('#'+id);
  const mapDialog=get('ec-map-dialog'), importEl=suffix=>get('ec-map-'+suffix);
  const importKey=value=>String(value||'').normalize('NFKC').toLowerCase().replace(/[^\p{L}\p{N}]/gu,'');
  let importSession=null,importSequence=0;
  function splitImportRows(text,delimiter,tailMode=false,probe=false){
    const values=[];let fields=[],cell='',quoted=false,closed=false,start=true,total=0;
    function pushField(){fields.push(cell);cell='';start=true;closed=false;if(++total>400000||fields.length>140)throw Error('This file has too many columns or cells. Split it into smaller files.');}
    function pushRow(){pushField();if(fields.some(value=>value.trim()!==''))values.push(fields);fields=[];}
    text=text.replace(/^\uFEFF/,'');
    for(let pos=0;pos<text.length;pos++){
      const ch=text[pos];
      if(quoted){
        if(ch==='"'&&text[pos+1]==='"'){cell+='"';pos++;}
        else if(ch==='"'){quoted=false;closed=true;}else cell+=ch;
      }else if(ch===delimiter&&(!tailMode||fields.length===0)){pushField();}
      else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&text[pos+1]==='\n')pos++;pushRow();if(probe&&values.length>=12)return values;}
      else if(start&&ch==='"'&&(!tailMode||fields.length===0)){quoted=true;start=false;}
      else if(closed){if(ch!==' '&&ch!=='\t')throw Error('Unexpected text after a quoted field. Check the separator and quotes.');}
      else{cell+=ch;start=false;}
    }
    if(quoted)throw Error('A quoted field is missing its closing quote.');
    if(cell||fields.length||closed)pushRow();
    return values;
  }
  function importAliases(adapter){
    const aliases=new Map();
    for(const field of adapter.fields)for(const label of [field.key,field.label,...(field.aliases||[])])if(label&&!aliases.has(importKey(label)))aliases.set(importKey(label),field.key);
    return aliases;
  }
  function detectImportDelimiter(raw,adapter){
    const aliases=importAliases(adapter);let best=',',score=0;
    for(const candidate of [',','\t',';',':','|']){
      let rows;try{rows=splitImportRows(raw,candidate,false,true);}catch{continue;}
      if(!rows.length)continue;
      const sizes=new Map();for(const row of rows)sizes.set(row.length,(sizes.get(row.length)||0)+1);
      for(const [width,count] of sizes){if(width<2)continue;const hits=rows[0].filter(text=>aliases.has(importKey(text))).length;
        const rank=hits*6+count/rows.length*3+Math.min(width,16)/16;if(rank>score){score=rank;best=candidate;}}
    }return best;
  }
  function visibleName(separator){return ({',':'Comma',';':'Semicolon',':':'Colon','\t':'Tab','|':'Pipe'})[separator]||'Custom';}
  function importFeedback(message,lines=[]){
    importEl('status').textContent=message;const nodes=lines.slice(0,40).map(line=>{const item=document.createElement('li');item.textContent=String(line);return item;});
    if(lines.length>40){const item=document.createElement('li');item.textContent=(lines.length-40)+' more messages';nodes.push(item);}
    importEl('issues').replaceChildren(...nodes);
  }
  function refreshImportPreview(){
    const session=importSession;if(!session||session.busy)return;session.plan=null;importEl('apply').disabled=true;
    const {adapter,mapping,records,header,type}=session,rows=header?records.slice(1):records;
    const fieldByKey=new Map(adapter.fields.map(field=>[field.key,field]));
    const hiddenFields=new Set();
    if(ecPrivacyEnabled())mapping.forEach((key,index)=>{
      const field=fieldByKey.get(key);
      if(field?.sensitive||(!field&&type!=='addresses'))hiddenFields.add(index);
    });
    importEl('sample').replaceChildren(...rows.slice(0,6).map((values,index)=>{
      const tr=document.createElement('tr'),num=document.createElement('td');num.textContent=String(index+(header?2:1));tr.append(num);
      for(let col=0;col<mapping.length;col++){const td=document.createElement('td'),text=values[col]||'';td.textContent=hiddenFields.has(col)&&text?'••••':text||'—';tr.append(td);}return tr;}));
    if(!rows.length){importFeedback(records.length?'There are no data rows. Uncheck the header option if the first row is a record.':'Choose a file or paste rows to begin.');return;}
    const active=mapping.filter(Boolean);
    if(!active.length){importFeedback('Choose which columns to import.');return;}
    if(new Set(active).size!==active.length){importFeedback('Each destination field can be mapped only once.');return;}
    try{
      const mapped=rows.map((values,index)=>{const record=Object.create(null);Object.defineProperty(record,'__row',{value:index+(header?2:1)});mapping.forEach((key,col)=>{if(key)record[key]=values[col]??'';});return record;});
      const plan=adapter.inspect(mapped,{mapping:[...mapping],separator:session.separator,header});
      const blocked=Boolean(plan.issues?.length);session.plan=!blocked&&plan.actions>0?plan:null;
      importFeedback(plan.summary||'Review the mapped rows.',[...(plan.issues||[]),...(plan.notes||[])]);
      importEl('apply').disabled=!session.plan;importEl('apply').textContent=plan.actions>0?'Import ('+plan.actions.toLocaleString()+')':'Import';
    }catch(error){importFeedback(error.message||'Could not preview the import.');}
  }
  function parseImportSource(keepMapping=false){
    const session=importSession;if(!session||session.busy)return;
    const privateImport=ecPrivacyEnabled()&&session.type!=='addresses';
    const {adapter,raw}=session;session.plan=null;importEl('apply').disabled=true;importEl('apply').textContent='Import';
    const choice=importEl('delimiter').value;importEl('custom-box').hidden=choice!=='custom';
    try{
      if(raw.length>4*1024*1024)throw Error('Choose a file under 4 MiB, or import it in smaller batches.');
      const sep=choice==='auto'?detectImportDelimiter(raw,adapter):choice==='tab'?'\t':choice==='custom'?importEl('custom').value:choice;
      if(sep.length!==1||/["\r\n]/.test(sep))throw Error('Enter one separator character other than a quote or newline.');
      session.separator=sep;session.records=splitImportRows(raw,sep,Boolean(session.preset));
      const rows=session.records,first=rows[0]||[],aliases=importAliases(adapter);
      if(!session.manualHeader){
        const names=first.filter(text=>text.trim());
        const matches=text=>aliases.has(importKey(text));
        const recognized=names.filter(matches).length;
        importEl('header').checked=recognized>0&&(recognized>=2||recognized===names.length);
      }
      session.header=importEl('header').checked;
      const width=rows.reduce((max,row)=>Math.max(max,row.length),0);
      const labels=Array.from({length:width},(_,index)=>session.header?first[index]||'Column '+(index+1):'Column '+(index+1));
      const signature=JSON.stringify([sep,session.header,session.header?labels:width]);
      if(!keepMapping||session.signature!==signature){
        session.mapping=labels.map((label,index)=>session.preset?adapter.preset.mapping[index]||'':session.header?aliases.get(importKey(label))||'':'');
      }
      session.signature=signature;
      const heading=document.createElement('tr');const number=document.createElement('th');number.textContent='#';heading.append(number);
      let hiddenHeaders=false;
      labels.forEach((label,index)=>{
        const th=document.createElement('th');th.scope='col';
        const unconfirmed=privateImport&&session.header&&!session.manualHeader&&!aliases.has(importKey(label));
        if(unconfirmed)hiddenHeaders=true;
        const concealHeader=privateImport&&(/(?:\d[ -]?){10,}/.test(label)||unconfirmed);
        th.textContent=concealHeader?'Column '+(index+1):label;
        const select=document.createElement('select');select.dataset.mapColumn=String(index);select.setAttribute('aria-label','Map column '+(index+1));
        select.append(new Option('Skip column',''));for(const field of adapter.fields)select.append(new Option(field.label,field.key));
        select.value=session.mapping[index];th.append(select);heading.append(th);
      });
      importEl('headings').replaceChildren(heading);importEl('table-section').hidden=!rows.length;importEl('reveal').hidden=!hiddenHeaders;
      importEl('source').textContent=rows.length?(session.filename?session.filename+' · ':'')+visibleName(sep)+(session.preset?' (split once)':'')+' · '+Math.max(0,rows.length-(session.header?1:0)).toLocaleString()+' data rows · preview up to 6':'';
      refreshImportPreview();
    }catch(error){session.records=[];importEl('source').textContent=session.filename||'';importEl('table-section').hidden=true;importEl('sample').replaceChildren();importFeedback(error.message);}
  }
  function closeMappedImport(force=false){
    if(importSession?.busy&&!force)return;
    importSequence++;importSession=null;if(mapDialog.open)mapDialog.close();
    for(const key of ['text','file','custom'])importEl(key).value='';
    for(const key of ['headings','sample','issues'])importEl(key).replaceChildren();
    importEl('source').textContent='';importEl('status').textContent='';
  }
  function openMappedImport(type){
    const adapter=mappedImportAdapters[type];if(!adapter||mapDialog.open)return;
    importSession={type,adapter,busy:false,raw:'',filename:'',records:[],mapping:[],manualHeader:false,preset:false,signature:''};
    importEl('title').textContent=adapter.title();importEl('preset').hidden=!adapter.preset;
    mapDialog.dataset.sourceType=type;
    if(adapter.preset)importEl('preset').textContent='Use '+adapter.preset.label;
    importEl('delimiter').value='auto';importEl('header').checked=true;
    for(const key of ['cancel','close'])importEl(key).disabled=false;
    parseImportSource();mapDialog.showModal();
  }
  importEl('text').addEventListener('input',()=>{
    if(!importSession)return;importSequence++;importSession.raw=importEl('text').value;importSession.filename='';importEl('file').value='';parseImportSource(true);
  });
  async function readImportFile(file){
    const session=importSession;if(!session||session.busy||!file)return;const sequence=++importSequence;
    session.plan=null;session.raw='';session.filename=file.name;importEl('text').value='';parseImportSource();
    if(file.size>4*1024*1024){importFeedback('Choose a file under 4 MiB, or import it in smaller batches.');return;}
    try{
      const raw=await file.text();if(sequence!==importSequence||importSession!==session)return;
      session.raw=raw;importEl('text').value=raw;parseImportSource();
    }
    catch{if(sequence===importSequence)importFeedback('Could not read the file. Paste its contents instead.');}
  }
  importEl('file').addEventListener('change',()=>readImportFile(importEl('file').files[0]));
  mapDialog.addEventListener('dragover',event=>{if(event.dataTransfer.types.includes('Files'))event.preventDefault();});
  mapDialog.addEventListener('drop',event=>{
    if(!event.dataTransfer.files.length)return;event.preventDefault();if(importSession?.busy)return;
    if(event.dataTransfer.files.length!==1){importFeedback('Drop one file at a time.');return;}
    importEl('file').value='';readImportFile(event.dataTransfer.files[0]);
  });
  for(const key of ['delimiter','custom','header'])importEl(key).addEventListener(key==='custom'?'input':'change',()=>{
    if(!importSession)return;importSession.preset=false;if(key==='header')importSession.manualHeader=true;parseImportSource();
  });
  importEl('preset').addEventListener('click',()=>{const session=importSession;if(!session?.adapter.preset)return;session.preset=true;session.manualHeader=true;importEl('header').checked=false;importEl('delimiter').value=session.adapter.preset.separator;parseImportSource();});
  importEl('reveal').addEventListener('click',()=>{if(importSession){importSession.manualHeader=true;parseImportSource(true);}});
  root.addEventListener('ec:privacy-change',()=>parseImportSource(true));
  importEl('headings').addEventListener('change',event=>{if(!importSession||!event.target.matches('[data-map-column]'))return;importSession.mapping[Number(event.target.dataset.mapColumn)]=event.target.value;refreshImportPreview();});
  for(const key of ['close','cancel'])importEl(key).addEventListener('click',()=>closeMappedImport());
  mapDialog.addEventListener('cancel',event=>{event.preventDefault();closeMappedImport();});
  importEl('apply').addEventListener('click',async()=>{
    const session=importSession;if(!session?.plan||session.busy)return;
    const plan=session.plan;session.busy=true;mapDialog.querySelector('.ec-map-body').inert=true;
    for(const key of ['apply','cancel','close'])importEl(key).disabled=true;
    try{await plan.apply();if(session===importSession)closeMappedImport(true);}
    catch(error){if(session===importSession){session.busy=false;refreshImportPreview();importFeedback(error.message||'Could not import. Review the rows and retry.');}}
    finally{mapDialog.querySelector('.ec-map-body').inert=false;if(session===importSession){session.busy=false;importEl('cancel').disabled=false;importEl('close').disabled=false;}}
  });
  const people=[{id:'p-ros',name:'Rosalyn'},{id:'p-ver',name:'Veronica'}];
  const addresses=[
    {id:'ros-home',personId:'p-ros',person:'Rosalyn',label:'Home',recipients:['Rosalyn Example','Josh Example'],street:'123 Example Street',street2:'Unit 2',city:'Sample City',state:'CA',zip:'90210',country:'US',phone:'202-555-0110',email:'rosalyn@example.test',color:'home'},
    {id:'ros-office',personId:'p-ros',person:'Rosalyn',label:'Office',firstName:'Rosalyn',lastName:'Example',street:'47 Sample Road',street2:'Suite 5',city:'Sample City',state:'CA',zip:'90211',country:'US',phone:'202-555-0110',email:'rosalyn@example.test',color:'office'},
    {id:'ver-home',personId:'p-ver',person:'Veronica',label:'Home',firstName:'Veronica',lastName:'Sample',street:'900 Example Avenue',street2:'',city:'Example City',state:'OR',zip:'97035',country:'US',phone:'202-555-0111',email:'veronica@example.test',color:'veronica'}
  ];
  const ignoredAddressKeys=new Set();
  const directoryAddresses=()=>addresses.filter(address=>!ignoredAddressKeys.has(address.id));
  const shippingGroupKey=parentId=>'shipping-group:'+parentId;
  const allShippingDestinations=()=>addresses.flatMap(shippingDestinationOptions);
  const directoryDestinations=()=>directoryAddresses().flatMap(shippingDestinationOptions);
  const activeAddresses=()=>directoryDestinations().filter(address=>address.active!==false);
  const isDestinationChosen=(list,address)=>list.chosen.has(address.id)||list.chosen.has(shippingGroupKey(address.parentAddressId));
  const unitAtEnd=/(?:^|[\s,])(?:(?:suite|ste|apartment|apt|unit)\.?(?:\s*#\s*|\s+)|#\s*)([a-z\d][\w/-]*)$/i;
  function streetLines(address,secondOverride){
    const street=address?.street||'',street2=address?.street2||'';
    const unit=street2.trim().match(unitAtEnd);
    if(!street.trim()||!unit||unit.index!==0)return {street,street2};
    // A profile can explicitly assign its own unit to the second field.
    const override=String(secondOverride||'').trim().match(unitAtEnd);
    if(override?.index===0)return {street,street2};
    const present=street.trimEnd().match(unitAtEnd);
    if(present&&present[1].toLowerCase()!==unit[1].toLowerCase())return {street,street2};
    return {street:present?street:street.trimEnd()+' '+street2.trim(),street2:''};
  }
  const streetAndUnit=address=>Object.values(streetLines(address)).filter(Boolean).join(', ');
  const addressName=address=>address.label||streetAndUnit(address);
  function shippingUnitSortKey(address){
    const unit=(address.street2||'').replace(/^(?:suite|ste|apartment|apt|unit)\.?\s*#?\s*|^#\s*/i,'').trim();
    if(unit)return unit;
    return (address.street||'').match(/(?:\b(?:suite|ste|apartment|apt|unit)\.?\s*#?\s*|#\s*)([a-z\d][\w/-]*)$/i)?.[1]||'';
  }
  function compareShippingDestinations(left,right,descending=false){
    const a=shippingUnitSortKey(left),b=shippingUnitSortKey(right);
    if(!!a!==!!b)return a?-1:1;
    const order=nameCollator.compare(a,b);
    if(order)return descending?-order:order;
    for(const field of ['street','street2','city','state','zip','country']){
      const difference=nameCollator.compare(left[field]||'',right[field]||'');
      if(difference)return difference;
    }
    return 0;
  }
  function shippingVariantLookup(parent){
    const ids=new Set(),authored=new Map();
    for(const suite of parent.shippingVariants||[]){
      ids.add(suite.id);
      if(suite.authored){
        const key=destinationLocationKey(suite);
        if(!authored.has(key))authored.set(key,suite.id);
      }
    }
    return observation=>{
      if(!observation?.variantId)return '';
      const parcelId=parent.id+':variant:'+observation.variantId;
      if(!authored.size||ids.has(parcelId))return parcelId;
      return authored.get(destinationLocationKey(observation))||parcelId;
    };
  }
  function destinationLocationKey(location){
    return JSON.stringify(Object.values(comparableLocation(location)).map(normalized));
  }
  function comparableLocation(source){
    const location=shippingLocation(source);
    return {...location,...streetLines(location)};
  }
  function shippingLocation(source){
    const clean=value=>String(value??'').trim().replace(/\s+/g,' ');
    const location=Object.fromEntries(['street','street2','city','state','zip','country'].map(field=>[field,clean(source[field])]));
    if(!isUSAddress(location.country))return location;
    const unitPattern=/^(suite|ste|apartment|apt|unit)\.?(?:\s*#\s*|\s+)([a-z]|\d[\w/-]*|[a-z]+\d[\w/-]*)$/i;
    const unitLabel=match=>{
      const kind=match[1].toLowerCase();
      return (kind==='suite'||kind==='ste'?'Suite':kind==='unit'?'Unit':'Apt')+' '+match[2];
    };
    const embedded=location.street.match(/^(.*?\S)[,\s]+((?:suite|ste|apartment|apt|unit)\.?(?:\s*#\s*|\s+)[\w/-]+)$/i);
    const embeddedUnit=embedded?.[2].match(unitPattern),separateUnit=location.street2.match(unitPattern);
    if(embeddedUnit&&(!location.street2||separateUnit&&unitLabel(separateUnit).toLowerCase()===unitLabel(embeddedUnit).toLowerCase())){
      location.street=embedded[1];
      location.street2=unitLabel(embeddedUnit);
    }else if(separateUnit)location.street2=unitLabel(separateUnit);
    function displayCase(value){
      if(value!==value.toUpperCase()&&value!==value.toLowerCase())return value;
      return value.replace(/[A-Za-z]+/g,(word,index)=>{
        if(/^(?:[NSEW]|NE|NW|SE|SW|PO|US)$/i.test(word))return word.toUpperCase();
        if(index&&/\d/.test(value[index-1]))return /^(?:st|nd|rd|th)$/i.test(word)?word.toLowerCase():word;
        return word[0].toUpperCase()+word.slice(1).toLowerCase();
      });
    }
    location.street=displayCase(location.street);
    location.city=displayCase(location.city);
    if(/^[A-Za-z]{2}$/.test(location.state))location.state=location.state.toUpperCase();
    location.country='US';
    const {base,extension}=splitAddressPostal(location.country,location.zip);
    location.zip=joinAddressPostal(base,extension);
    return location;
  }
  function ensureShippingVariant(parent,observation){
    const id=shippingVariantLookup(parent)(observation);
    if(!id)throw Error('Refresh workspace to identify this shipping destination.');
    const variants=parent.shippingVariants??=[];
    let stored=variants.find(item=>item.id===id);
    if(!stored){
      stored={id,...shippingLocation(observation),enabled:false,recipients:[]};
      variants.push(stored);
    }
    return stored;
  }
  function shippingDestinationOptions(parent){
    const available=parent.active!==false&&!parent.parcelMissing;
    const mainEnabled=parent.mainShippingEnabled!==false;
    const main={...parent,parentAddressId:parent.id,shippingEnabled:mainEnabled,active:available&&mainEnabled};
    const children=(parent.shippingVariants||[]).map(suite=>{
      const label=[parent.label,suite.street2||suite.street].filter(Boolean).join(' · ');
      return {...parent,...suite,id:suite.id,parentAddressId:parent.id,label,
        recipients:suite.recipients?.length?suite.recipients:addressRecipients(parent),
        email:suite.email||parent.email,phone:suite.phone||parent.phone,shippingEnabled:suite.enabled===true,active:available&&suite.enabled===true};
    });
    return [main,...children];
  }
  function addressRecipients(address){
    if(Array.isArray(address.recipients))return address.recipients;
    const legacy=formatRecipientParts(address.firstName,address.lastName);
    return legacy?[legacy]:[];
  }
  function parseRecipientName(value){
    const text=String(value??'').trim(),parts=text.split('|');
    if(parts.length===2){
      const [firstName,lastName]=parts.map(part=>part.trim());
      const fullName=[firstName,lastName].filter(Boolean).join(' ');
      return {firstName,lastName,fullName,explicit:true,valid:!!fullName};
    }
    const match=text.match(/^(\S+)\s+([\s\S]+)$/);
    return {firstName:match?match[1]:text,lastName:match?match[2]:'',fullName:text,explicit:parts.length>1,valid:!!text&&parts.length===1};
  }
  function formatRecipientParts(first,last){
    const firstName=String(first??'').trim(),lastName=String(last??'').trim();
    const fullName=[firstName,lastName].filter(Boolean).join(' '),guess=parseRecipientName(fullName);
    return guess.firstName===firstName&&guess.lastName===lastName?fullName:(firstName+' | '+lastName).trim();
  }
  function recipientNameKey(name){
    return normalized(parseRecipientName(name).fullName).replace(/\s+/g,' ');
  }
  function mergeRecipientName(names,name,explicitOverride=false){
    const key=recipientNameKey(name),index=names.findIndex(other=>recipientNameKey(other)===key);
    if(index<0)return [...names,name];
    if(!explicitOverride&&!parseRecipientName(name).explicit||names[index]===name)return names;
    return names.map((other,i)=>i===index?name:other);
  }
  function recipientFields(address,sequence){
    const recipients=addressRecipients(address);
    const name=recipients[(sequence-1)%recipients.length]||'';
    const parsed=parseRecipientName(name);
    if(!parsed.explicit&&name&&name===[address.firstName,address.lastName].filter(Boolean).join(' '))
      return {shippingName:name,firstName:address.firstName||'',lastName:address.lastName||''};
    return {shippingName:parsed.fullName,firstName:parsed.firstName,lastName:parsed.lastName};
  }
  function setAddressDetails(address,changes){
    for(const field of ['recipients','active','email','phone']){
      if(Object.hasOwn(changes,field))address[field]=changes[field];
    }
  }
  const masterCards=['Card 01','Card 02','Card 03'];
  const cardGroups=[{name:'Employee cards',cards:['Card 01','Card 02']},{name:'Virtual cards',cards:['Card 03']}];
  const accountGroups=[],removedAccountKeys=new Set();
  let parcelMailSearchUrl=()=>'';
  const managedRetailAccount=account=>!account.parcelKey||account.retailManaged===true;
  const supportedRetailers=['Target','Walmart','Pokémon Center','Riot'];
  const poolMemberships=new WeakMap();
  function poolMembership(ids){
    if(!Array.isArray(ids))return null;
    // Each edit replaces this array, so saved lists and earlier selections keep their own index.
    let members=poolMemberships.get(ids);
    if(!members){members=new Set(ids);poolMemberships.set(ids,members);}
    return members;
  }
  function accountMatchesPool(account,list,retained=false){
    if(account.retailer!==(list.retailer||'Target'))return false;
    const members=poolMembership(list.accountGroupIds);
    return !members||(account.groupIds||[]).some(id=>members.has(id))||
      (retained&&account.formerParcelGroups?.some(id=>members.has(id)));
  }
  const profileColumns=[
    ['nickname','Profile nickname'],['email','Account email'],['login','Retail username'],['card','Card'],['phone','Phone'],
    ['firstName','Shipping first name'],['lastName','Shipping last name'],['street','Shipping street'],['street2','Shipping street 2'],['street3','Shipping street 3'],['city','Shipping city'],['state','Shipping state'],['zip','Shipping ZIP'],['country','Shipping country'],
    ['billingFirstName','Billing first name'],['billingLastName','Billing last name'],['billingPhone','Billing phone'],['billingStreet','Billing street'],['billingStreet2','Billing street 2'],['billingStreet3','Billing street 3'],['billingCity','Billing city'],['billingState','Billing state'],['billingZip','Billing ZIP'],['billingCountry','Billing country'],
    ['group','Profile group'],['notes','Notes'],['onlyCheckoutOnce','Only checkout once'],['matchNameOnCardAndAddress','Match name on card and address']
  ];
  const readOnlyFields=new Set(['email','login','card']);
  const columnHelp={
    login:'Username for the assigned retail account, if different from its email.',
    firstName:'The shipping recipient saved on this suite or address, unless edited in this profile.',
    lastName:'The shipping recipient saved on this suite or address, unless edited in this profile.',
    email:'Email from the assigned retail account. Edit the account in Retail Accounts.'
  };
  const editableFields=new Set(profileColumns.map(([key])=>key).filter(key=>!readOnlyFields.has(key)));
  const nameCollator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
  const flagFields=new Set(['onlyCheckoutOnce','matchNameOnCardAndAddress']);
  const paymentColumns=[['holder','Cardholder'],['brand','Card type'],['number','Card number'],['month','Expiry month'],['year','Expiry year'],['cvv','Security code (CVV/CVC)'],['paymentType','Payment method']];
  const fieldLabels=new Map([...profileColumns,...paymentColumns,['sameBillingAndShippingAddress','Same billing and shipping flag']]);
  const profileFieldKeys=['nickname','group','email','login','card','notes',...flagFields];
  const billingFieldKeys=profileColumns.map(([key])=>key).filter(key=>key.startsWith('billing'));
  const shippingFieldKeys=profileColumns.map(([key])=>key).filter(key=>!profileFieldKeys.includes(key)&&!billingFieldKeys.includes(key));
  const fieldSections=[['Profile',profileFieldKeys],['Shipping',shippingFieldKeys],['Billing',[...billingFieldKeys,'sameBillingAndShippingAddress']],['Payment',paymentColumns.map(([key])=>key)]];
  let visibleColumns=[];
  const fittedWidths=new Map(),customWidths=new Map();
  let columnEditor=null;
  const sheetMeasure=document.createElement('canvas').getContext('2d');
  const masterAccounts=Array.from({length:1000},(_,i)=>({id:i+1,retailer:'Target',email:'account'+String(i+1).padStart(4,'0')+'@example.test',username:'',password:''}));
  for(const retailer of ['Walmart','Riot'])masterAccounts.push({id:masterAccounts.length+1,retailer,email:retailer.toLowerCase()+'@example.test',username:retailer==='Riot'?'ExamplePlayer':'',password:''});
  const cardLastFour=new Map(),cardNetworks=new Map(),cardsWithSecurityCode=new Set();
  const networkPatterns={
    "Visa":"^4[0-9]{12}(?:[0-9]{3}|[0-9]{5,6})?$",
    "Mastercard":"^(?:5[1-5][0-9]{14}|(?:222[1-9]|22[3-9][0-9]|2[3-6][0-9]{2}|27[01][0-9]|2720)[0-9]{12})$",
    "American Express":"^3[47][0-9]{13}$",
    "Discover":"^(?:6011[0-9]{12}|64[4-9][0-9]{13}|65[0-9]{14})(?:[0-9]{3})?$",
    "JCB":"^(?:352[89]|35[3-8][0-9])[0-9]{12,15}$",
    "Diners Club":"^(?:30[0-5][0-9]{11}|3[689][0-9]{12})(?:[0-9]{2}|[0-9]{5})?$"
  };
  const networkRules=Object.entries(networkPatterns).map(([network,pattern])=>[network,new RegExp(pattern)]);
  function cardNumberIssue(digits){
    if(!/^[0-9]{12,19}$/.test(digits))return 'Enter a card number with 12–19 digits.';
    let sum=0,doubleDigit=false;
    for(let index=digits.length-1;index>=0;index--){
      let digit=Number(digits[index]);
      if(doubleDigit){digit*=2;if(digit>9)digit-=9;}
      sum+=digit;doubleDigit=!doubleDigit;
    }
    return sum>0&&sum%10===0?'':'The card number failed its checksum. Check for a mistyped digit.';
  }
  function detectedNetwork(number){
    return networkRules.find(([,pattern])=>pattern.test(number||''))?.[0]||'';
  }
  function displayNetwork(brand,number){
    const value=String(brand||'').trim();
    return value&&value.length<=80&&!/(?:[0-9][ -]?){12}/.test(value)?value:detectedNetwork(number);
  }
  const groupsForCard=new Map();
  function cardGroupNames(label){return groupsForCard.get(label)||[];}
  function cardSheetCaption(label,selectedGroups){
    const memberships=cardGroupNames(label),chosen=memberships.filter(name=>selectedGroups.has(name));
    return [...(chosen.length?chosen:memberships),label].join(' · ');
  }
  const cardBilling=new Map();
  const exportLedger=[];
  const sampleCardDetails=new Map();
  let resolveCheckout=async ids=>new Map(ids.filter(id=>sampleCardDetails.has(id)).map(id=>[id,sampleCardDetails.get(id)]));
  function noteCardChange(id,details){
    if('cvv' in details){if(details.cvv)cardsWithSecurityCode.add(id);else cardsWithSecurityCode.delete(id);}
    invalidateDirectorySecrets();
  }
  let stageImportedCard=(id,details)=>{
    noteCardChange(id,details);sampleCardDetails.set(id,{...sampleCardDetails.get(id),...details});
  },saveCoordinatorChange=()=>{};
  let finishInventoryImport=async()=>{};
  async function getExportCardDetails(ids){
    const cards=await resolveCheckout([...new Set(ids)]);
    return new Map([...cards].map(([id,card])=>[id,{...card,brand:displayNetwork(card.brand,card.number)}]));
  }
  function recordProfileExport(summary){
    if(!summary||!Array.isArray(summary.profiles))return;
    const {id,eventId,eventName,listId,listName,bot,format,at}=summary;
    if(!id||exportLedger.some(item=>item.id===id))return;
    exportLedger.push({id,eventId,eventName,listId,listName,bot,format,at,profiles:summary.profiles.map(({accountId,addressId,destinationId,cardId,nickname})=>({accountId,addressId,...(destinationId?{destinationId}:{}),cardId,nickname}))});
    saveCoordinatorChange();
  }
  let localAccountSerial=masterAccounts.length;
  const nextAccountId=()=>++localAccountSerial;
  const cardIds=new Map(masterCards.map((card,index)=>[card,'card-'+(index+1)]));
  let nextCardId=cardIds.size+1;
  const freshSeed=()=>Math.floor(Math.random()*4294967296);
  const formatNames={sniped:'Sniped',shikari:'Shikari',aycd:'AYCD'};
  const formatLabel=format=>formatNames[format];
  const nextListId=()=>'list-'+Array.from(crypto.getRandomValues(new Uint8Array(16)),byte=>byte.toString(16).padStart(2,'0')).join('');
  const makeList=(id,name,format)=>({id,name,format,sort:1,rows:[],chosen:new Set(),cards:new Set(),useUngroupedCards:false,disabledColumns:new Set(),pool:true,mode:'even',basis:'address',quantity:0,nicknameCounters:{},names:new Map(),edits:new Map(),seed:freshSeed()});
  const makeEvent=(id,name)=>({id,name,sharing:false,lists:[],selectedListId:null});
  const events=[];
  let activeEvent=null,eventSerial=1;
  const esc=v=>String(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const currentEvent=()=>events.find(e=>e.id===activeEvent)||null;
  const currentList=()=>{const event=currentEvent();return event?.lists.find(list=>list.id===event.selectedListId)||null;};
  const billingFields={firstName:'billingFirstName',lastName:'billingLastName',phone:'billingPhone',street:'billingStreet',street2:'billingStreet2',street3:'billingStreet3',city:'billingCity',state:'billingState',zip:'billingZip',country:'billingCountry'};
  function profilePhone(value,country){
    const original=String(value??''),local=original.trim().match(/^\+1[ .-]*([\d(][\d ().-]*)$/);
    return local&&isUSAddress(country)&&local[1].replace(/\D/g,'').length===10?local[1].trim():original;
  }
  function syncProfileBilling(row,edits){
    const shippingLines=streetLines(row.shippingSource,edits.street2);
    row.street=edits.street??shippingLines.street;row.street2=edits.street2??shippingLines.street2;
    const billingLines=streetLines(row.billingSource,edits.billingStreet2);
    for(const [shipping,billing] of Object.entries(billingFields)){
      const source=edits.matchShippingBilling&&!edits.defaultBillingFields?.includes(billing)?null:row.billingSource;
      const separateStreet=(shipping==='street2'||shipping==='street3')&&source?.street&&normalized(billingLines.street)!==normalized(row.street);
      const value=source&&(shipping==='street'||shipping==='street2'?billingLines[shipping]:source[shipping]);
      row[billing]=edits[billing]??(value||(separateStreet?'':row[shipping])||'');
    }
    if(!row.billingCountry)row.billingCountry=row.country||'US';
    row.zip=profilePostal(row.zip,row.country);
    row.billingZip=profilePostal(row.billingZip,row.billingCountry);
    row.billingName=[row.billingFirstName,row.billingLastName].filter(Boolean).join(' ');
  }
  function withDraftEdits(saved,draft){
    if(!draft)return saved||{};
    const overrides={...saved};
    if(draft.matchShippingBilling){
      for(const field of Object.values(billingFields))delete overrides[field];
      delete overrides.defaultBillingFields;
    }
    return Object.assign(overrides,draft);
  }
  // Undo stores override changes, so a restore or undo always uses the current linked sources.
  const sheetHistories=new WeakMap();
  let typingAction=null;
  function sheetHistory(list){
    let history=sheetHistories.get(list);
    if(!history){history={undo:[],redo:[],draftKeys:new Set()};sheetHistories.set(list,history);}
    return history;
  }
  const sameEditValue=(a,b)=>a===b||(Array.isArray(a)&&Array.isArray(b)&&a.length===b.length&&a.every((v,i)=>v===b[i]));
  function changedOverrideFields(before,after){
    return [...new Set([...Object.keys(before),...Object.keys(after)])].filter(field=>!sameEditValue(before[field],after[field]));
  }
  function setSheetOverrides(list,key,overrides){
    if(Object.keys(overrides).length)list.edits.set(key,overrides);
    else list.edits.delete(key);
  }
  function overrideStep(current,change,direction){
    if(!change)return current;
    const edits={...current},snapshot=direction==='undo'?change.before:change.after;
    for(const field of changedOverrideFields(change.before,change.after)){
      if(Object.hasOwn(snapshot,field))edits[field]=snapshot[field];else delete edits[field];
    }
    return edits;
  }
  function trackSheetEdit(list,keys,edit,input=null){
    const history=sheetHistory(list);
    const merging=input&&typingAction?.input===input&&typingAction.action===history.undo.at(-1);
    const action=merging?typingAction.action:new Map();
    const prior=new Map([...keys].map(key=>[key,list.edits.get(key)||{}]));
    edit();
    let changed=false;
    for(const [key,old] of prior){
      const after=list.edits.get(key)||{};
      if(!changedOverrideFields(old,after).length)continue;
      changed=true;
      const before=action.get(key)?.before||old;
      if(changedOverrideFields(before,after).length)action.set(key,{before,after});
      else action.delete(key);
      if(typeof key==='string')history.draftKeys.add(key);
    }
    if(!changed)return;
    if(!merging&&action.size){
      history.undo.push(action);
      if(history.undo.length>100)history.undo.shift();
    }else if(merging&&!action.size)history.undo.pop();
    history.redo.length=0;
    typingAction=input&&action.size?{input,action}:null;
    updateSheetHistoryControls();
  }
  function promoteSheetHistory(list,draftKey,accountId){
    const history=sheetHistories.get(list);
    if(!history?.draftKeys.delete(draftKey))return;
    // A pending edit may have been undone before its account became available.
    // Rebind both stacks, preserving any overrides the account already owned.
    for(const direction of ['undo','redo']){
      const stack=history[direction];
      let account=list.edits.get(accountId)||{},draft=list.edits.get(draftKey)||{};
      for(let i=stack.length-1;i>=0;i--){
        const action=stack[i];
        if(!action.has(draftKey)&&!action.has(accountId))continue;
        const current=withDraftEdits(account,draft);
        account=overrideStep(account,action.get(accountId),direction);
        draft=overrideStep(draft,action.get(draftKey),direction);
        const next=withDraftEdits(account,draft);
        action.delete(draftKey);
        if(changedOverrideFields(current,next).length){
          action.set(accountId,direction==='undo'?{before:next,after:current}:{before:current,after:next});
        }else action.delete(accountId);
        if(!action.size)stack.splice(i,1);
      }
    }
  }
  function nextAddressNumber(counts,addressId){
    const number=(counts.get(addressId)||0)+1;
    counts.set(addressId,number);
    return number;
  }
  function buildRow(list,account,address,card,addressNumber,previewSuffix,sourceAccountId=account?.id??null){
    const key=account?account.id+':'+address.id:null;
    if(previewSuffix===undefined&&!list.names.has(key)){
      const number=(list.nicknameCounters[address.id]||0)+1;list.nicknameCounters[address.id]=number;
      list.names.set(key,String(number).padStart(2,'0'));
    }
    const suffix=previewSuffix??list.names.get(key),recipient=recipientFields(address,Number(suffix));
    // The stored slot keeps drafts, recipients and undo history tied to the same assignment.
    // Its display number follows the current allocation and can change when the list shrinks.
    const draftKey='draft:'+address.id+':'+suffix;
    const draft=list.edits.get(draftKey);
    const edits=withDraftEdits(list.edits.get(sourceAccountId),draft);
    if(account&&previewSuffix===undefined)promoteSheetHistory(list,draftKey,account.id);
    if(draft&&account&&previewSuffix===undefined){
      list.edits.set(account.id,edits);
      list.edits.delete(draftKey);
    }
    const row={accountId:account?.id??null,sourceAccountId,draftKey:previewSuffix===undefined?null:draftKey,addressId:address.id,parentAddressId:address.parentAddressId,color:address.color,
      nickname:[address.person,addressName(address),'#'+String(addressNumber).padStart(2,'0')].filter(Boolean).join(' '),card,
      phone:address.phone||'',...recipient,
      shippingSource:{street:address.street,street2:address.street2},street3:address.street3||'',city:address.city,state:address.state,zip:address.zip,country:address.country,
      billingSource:cardBilling.get(cardIds.get(card))||address.billing,group:'',notes:'',onlyCheckoutOnce:false,matchNameOnCardAndAddress:false,...edits};
    // Old sheet edits can contain contact details. The current assignment owns the profile email.
    row.accountEmail=(account?.email||'').trim();
    row.email=row.accountEmail;
    const username=(account?.username||'').trim();
    row.login=username.toLowerCase()===row.email.toLowerCase()?'':username;
    row.country=row.country||'US';
    row.phone=profilePhone(row.phone,row.country);
    if('firstName' in edits||'lastName' in edits)row.shippingName=[row.firstName,row.lastName].filter(Boolean).join(' ');
    syncProfileBilling(row,edits);
    return row;
  }
  function editSheetRow(list,row,update){
    const key=row.draftKey??row.accountId;
    if(key==null)return row;
    const edits={...list.edits.get(key)};
    update(edits);
    syncProfileBilling(row,row.draftKey?withDraftEdits(list.edits.get(row.sourceAccountId),edits):edits);
    list.edits.set(key,edits);
    return row;
  }
  function applySheetEdits(list,row,changes){
    return editSheetRow(list,row,edits=>{
      for(const [field,value] of Object.entries(changes)){
        if(!editableFields.has(field))continue;
        if(flagFields.has(field))edits[field]=value===true||/^(true|yes|1)$/i.test(String(value).trim());
        else if(field==='phone'||field==='billingPhone')edits[field]=profilePhone(value,field==='phone'?row.country:row.billingCountry);
        else edits[field]=value;
        row[field]=edits[field];
      }
      row.shippingName=[row.firstName,row.lastName].filter(Boolean).join(' ');
    });
  }
  function matchSheetBilling(list,row){
    return editSheetRow(list,row,edits=>{
      edits.matchShippingBilling=true;
      delete edits.defaultBillingFields;
      for(const field of Object.values(billingFields))delete edits[field];
    });
  }
  function refreshSourceRows(changedAddressId,onlyList){
    const byId=new Map(allShippingDestinations().map(address=>[address.id,address])),accountById=new Map(masterAccounts.map(account=>[account.id,account]));
    for(const list of onlyList?[onlyList]:events.flatMap(event=>event.lists)){
      const counts=new Map();
      list.rows=list.rows.map(row=>{
        const number=nextAddressNumber(counts,row.addressId);
        if(changedAddressId&&row.parentAddressId!==changedAddressId&&row.addressId!==changedAddressId)return row;
        const address=byId.get(row.addressId);const account=accountById.get(row.accountId);return address&&account?buildRow(list,account,address,row.card,number):row;
      });
    }
  }
  function otherListAccounts(event,list){
    return new Set(event.lists.filter(item=>item!==list).flatMap(item=>item.rows.map(row=>row.accountId)));
  }
  function available(event=currentEvent(),list=currentList()){
    if(!list?.pool||!event)return[];
    const sources=event.sharing?[list]:event.lists;
    const used=new Set(sources.flatMap(item=>item.rows.map(row=>row.accountId)));
    return masterAccounts.filter(account=>managedRetailAccount(account)&&accountMatchesPool(account,list)&&!used.has(account.id));
  }
  function seededRandom(seed){let value=seed>>>0;return()=>{value=(value+0x6d2b79f5)>>>0;let mixed=Math.imul(value^(value>>>15),1|value);mixed^=mixed+Math.imul(mixed^(mixed>>>7),61|mixed);return((mixed^(mixed>>>14))>>>0)/4294967296;};}
  function shuffle(values,random){const result=[...values];for(let i=result.length-1;i>0;i--){const j=Math.floor(random()*(i+1));[result[i],result[j]]=[result[j],result[i]];}return result;}
  let rememberTabView=()=>{};
  function storeView(){rememberTabView();if(window.openai?.setWidgetState)window.openai.setWidgetState({modelContent:{preview:'event-coordinator',eventId:activeEvent,listId:currentList()?.id||null},privateContent:null}).catch(()=>{});}
  function hierarchyStrip(kind){return get(kind==='event'?'ec-event-buttons':'ec-list-buttons');}
  function hierarchyRecords(kind){return kind==='event'?events:currentEvent()?.lists||[];}
  function renderStrip(kind,records,selectedId){
    const strip=hierarchyStrip(kind),key=JSON.stringify(records.map(item=>item.id));
    if(strip.dataset.structure!==key){
      strip.dataset.structure=key;
      strip.innerHTML=records.map(item=>{
        const name=esc(item.name),id=esc(item.id);
        return '<span class="ec-strip-item"><button type="button" role="tab" id="ec-tab-'+id+'" data-'+kind+'="'+id+'" aria-label="'+name+'" aria-controls="'+(kind==='event'?'ec-list-strip':'ec-list-panel')+'"><span class="ec-strip-name">'+name+'</span><span class="ec-count" aria-hidden="true"></span></button><span class="ec-strip-delete">'+deleteControl(kind,item.id,item.name)+'</span></span>';
      }).join('');
    }
    for(const [index,tab] of strip.querySelectorAll('[role=tab]').entries()){
      const item=records[index];
      if(tab.getAttribute('aria-label')!==item.name){
        tab.setAttribute('aria-label',item.name);tab.querySelector('.ec-strip-name').textContent=item.name;
        tab.parentElement.querySelector('.ec-strip-delete').innerHTML=deleteControl(kind,item.id,item.name);
      }
      const selected=tab.dataset[kind]===selectedId;
      tab.setAttribute('aria-selected',String(selected));tab.tabIndex=selected?0:-1;
    }
  }
  function displayedCount(list,options){
    return requestedProfiles(list)>list.rows.length&&options.some(address=>isDestinationChosen(list,address))?requestedProfiles(list):list.rows.length;
  }
  function updateHierarchyCounts(){
    const options=activeAddresses();
    for(const kind of ['event','list']){
      const records=hierarchyRecords(kind),tabs=hierarchyStrip(kind).querySelectorAll('[role=tab]');
      for(const [index,tab] of tabs.entries()){
        const item=records[index];if(!item)continue;
        const n=kind==='event'?item.lists.reduce((sum,list)=>sum+displayedCount(list,options),0):displayedCount(item,options);
        tab.querySelector('.ec-count').textContent=n.toLocaleString();
        tab.title=item.name+' · '+n.toLocaleString()+(n===1?' profile':' profiles');
      }
    }
  }
  function renderEvents(){
    const event=currentEvent();
    renderStrip('event',events,activeEvent);
    renderStrip('list',event?.lists||[],event?.selectedListId);
    updateHierarchyCounts();
  }
  function conflict(){
    const event=currentEvent(),seen=new Set(),overlap=new Set();
    for(const list of event.lists)for(const row of list.rows){
      if(seen.has(row.accountId))overlap.add(row.accountId);
      seen.add(row.accountId);
    }
    get('ec-conflict').hidden=event.sharing||overlap.size===0;
    get('ec-conflict').textContent=overlap.size+' accounts appear in multiple lists with account sharing turned off.';
  }
  // Overscan a few rows on either side so a swipe or a keyboard step stays continuous.
  const pickerWindowThreshold=100;
  function pickerWindow(id,markup,sync=()=>{},enabled=()=>true){
    const view=get(id),details=view.closest('details'),height=44,overscan=6;
    let rows=[],windowed=false,first=-1,last=-1,frame=0;
    function paint(force=false){
      if(!windowed){
        if(force){view.innerHTML=rows.map(markup).join('');sync();}
        return;
      }
      if(details&&!details.open){
        if(view.childElementCount)view.replaceChildren();
        first=last=-1;return;
      }
      const top=Math.min(Math.floor(view.scrollTop/height),Math.max(0,rows.length-1));
      const start=Math.max(0,top-overscan),end=Math.min(rows.length,top+Math.ceil((view.clientHeight||352)/height)+overscan+1);
      if(!force&&first===start&&last===end)return;
      const focus=document.activeElement,old=focus?.closest('[data-picker-row]');
      const focusIndex=old&&view.contains(old)?Number(old.dataset.pickerRow):-1;
      const focusClass=focus?.classList.contains('ec-remove-confirm')?'.ec-remove-confirm':'input,button';
      first=start;last=end;
      view.innerHTML='<div aria-hidden="true"></div>'+rows.slice(start,end).map((row,i)=>
        '<div class="ec-picker-row" data-picker-row="'+(i+start)+'">'+markup(row)+'</div>').join('')+'<div aria-hidden="true"></div>';
      view.firstElementChild.style.height=start*height+'px';
      view.lastElementChild.style.height=(rows.length-end)*height+'px';
      sync();
      if(focusIndex>=start&&focusIndex<end)view.querySelector('[data-picker-row="'+focusIndex+'"] '+focusClass)?.focus({preventScroll:true});
    }
    function schedule(){
      if(!frame&&windowed)frame=requestAnimationFrame(()=>{frame=0;paint();});
    }
    view.addEventListener('scroll',schedule,{passive:true});
    details?.addEventListener('toggle',()=>paint());
    new ResizeObserver(schedule).observe(view);
    view.addEventListener('keydown',event=>{
      const row=event.target.closest('[data-picker-row]');
      if(!windowed||!row||!['Tab','End','Home','ArrowDown','ArrowUp'].includes(event.key))return;
      const i=Number(row.dataset.pickerRow);
      const step=event.key==='ArrowUp'||event.key==='Tab'&&event.shiftKey?-1:1;
      let target=event.key==='End'?rows.length-1:event.key==='Home'?0:i+step;
      const direction=event.key==='End'?-1:event.key==='Home'?1:step;
      while(target>=0&&target<rows.length&&!enabled(rows[target]))target+=direction;
      if(target<0||target>=rows.length||event.key==='Tab'&&target>=first&&target<last)return;
      event.preventDefault();
      view.scrollTop=target*height;paint();
      view.querySelector('[data-picker-row="'+target+'"] :is(input,button):not([disabled],[hidden])')?.focus({preventScroll:true});
    });
    return {set(items,reset=false){
      rows=items;windowed=rows.length>pickerWindowThreshold;view.classList.toggle('ec-picker-window',windowed);
      if(reset)view.scrollTop=0;
      paint(true);
    },sync};
  }
  let addressPickerCounts=new Map(),addressPickerList=null,accountPoolIds=null;
  let previousAccountOptions=[],previousPoolRetailer='',cardBadgeSelection=new Set(),usingLooseCards=false,looseCount=0;
  function syncAddressPicker(){
    for(const box of get('ec-address-options').querySelectorAll('input')){
      const {parent,destinationGroup,address}=box.dataset;
      if(address){
        box.checked=!!addressPickerList&&(addressPickerList.chosen.has(address)||addressPickerList.chosen.has(shippingGroupKey(box.dataset.owner)));
      }else{
        const stats=addressPickerCounts.get(parent?'p:'+parent:'g:'+destinationGroup);
        box.checked=!!stats&&stats[0]===stats[1];box.indeterminate=!!stats&&stats[0]>0&&!box.checked;
      }
    }
  }
  function updateAddressPickerSelection(list,options){
    addressPickerList=list;addressPickerCounts=new Map();
    for(const address of options){
      const selected=Number(isDestinationChosen(list,address));
      for(const key of ['p:'+(address.personId||'__unassigned__'),'g:'+address.parentAddressId]){
        let stats=addressPickerCounts.get(key);
        if(!stats){stats=[0,0];addressPickerCounts.set(key,stats);}
        stats[0]+=selected;stats[1]++;
      }
    }
    addressPicker.sync();
  }
  function syncAccountPicker(){
    for(const box of get('ec-account-group-options').querySelectorAll('input')){
      box.checked=box.hasAttribute('data-pool-all')?accountPoolIds===null:!!accountPoolIds?.has(box.dataset.poolGroup);
    }
  }
  const addressPicker=pickerWindow('ec-address-options',row=>{
    const [attribute,id,label,depth,owner]=row;
    return '<label class="ec-option '+depth+'" title="'+esc(label)+'"><input type="checkbox" '+attribute+'="'+esc(id)+'"'+(owner?' data-owner="'+esc(owner)+'"':'')+' class="cursor-interaction"><span>'+esc(label)+'</span></label>';
  },syncAddressPicker);
  const addressChips=pickerWindow('ec-address-chips',row=>chip('address',...row));
  const accountPicker=pickerWindow('ec-account-group-options',group=>{
    const attr=group.id?'data-pool-group="'+esc(group.id)+'"':'data-pool-all';
    return '<label class="ec-option '+(group.id?'ec-child':'')+'" title="'+esc(group.name)+'"><input type="checkbox" '+attr+'><span>'+esc(group.name)+'</span></label>';
  },syncAccountPicker);
  const cardPicker=pickerWindow('ec-card-options',row=>cardBadge(row.attr,row.name,row.count),syncCardBadgeState,
    row=>row.count>0||(row.group?cardBadgeSelection.has(row.group):usingLooseCards));
  function counts(){
    const event=currentEvent(),list=currentList();if(!list)return;
    const pool=available(),count=list.rows.length;
    const retailer=list.retailer||'Target',savedAccounts=masterAccounts.filter(a=>managedRetailAccount(a)&&accountMatchesPool(a,list)).length;
    get('ec-list-retailer').value=retailer;get('ec-pool-title').textContent='Select retail accounts';get('ec-pool-caption').textContent='Use saved '+retailer+' accounts';
    const groups=accountGroups.filter(g=>g.retailer===retailer),ids=list.accountGroupIds;
    accountPoolIds=poolMembership(ids);
    if(previousPoolRetailer!==retailer||previousAccountOptions.length!==groups.length||
      groups.some((group,i)=>group.id!==previousAccountOptions[i].id||group.name!==previousAccountOptions[i].name)){
      const switched=previousPoolRetailer!==retailer;previousPoolRetailer=retailer;
      previousAccountOptions=groups.map(({id,name})=>({id,name}));
      accountPicker.set(groups.length?[{name:'All '+retailer},...previousAccountOptions]:[],switched);
    }else accountPicker.sync();
    const labels=ids==null?[retailer]:groups.filter(g=>accountPoolIds.has(g.id)).map(group=>group.name);
    const accountChips=get('ec-account-chips'),poolLabel=list.pool?labels.join(' + ')+(labels.length?' · ':'')+savedAccounts.toLocaleString()+' accounts':'';
    if(accountChips.dataset.poolLabel!==poolLabel){
      accountChips.dataset.poolLabel=poolLabel;accountChips.innerHTML=poolLabel?chip('pool',retailer,poolLabel,'ec-account-pool'):'';
    }
    updateHierarchyCounts();
    get('ec-account-count').textContent=count.toLocaleString()+' assigned here · '+pool.length.toLocaleString()+' unused '+(event.sharing?'in this list':'in this event');
    get('ec-select-unused').disabled=!pool.length;
    get('ec-basis').disabled=list.mode==='random';conflict();
    const missing=!list.pool?'Select a retail account pool.':!activeAddresses().some(address=>isDestinationChosen(list,address))?'Select an address.':!list.cards.size?'Select a card group.':list.quantity<1?'Set the number of profiles to start this list.':'No unused accounts available.';
    const shown=sheetRows.length,pending=shown-count;
    get('ec-empty').hidden=shown>0;get('ec-empty').textContent=Number(list.quantity)>100000?'A list supports up to 100,000 profiles.':missing;
    get('ec-live-status').textContent=pending?count.toLocaleString()+' assigned · '+pending.toLocaleString()+' pending':'Live sheet';
    get('ec-rendered-count').textContent=shown.toLocaleString()+(shown===1?' profile':' profiles');
    const exportButton=get('ec-open-export');
    exportButton.disabled=!shown;exportButton.title=shown?'':'Add profiles to export this list.';
  }
  function drawCardGroups(){
    groupsForCard.clear();
    for(const group of cardGroups)for(const label of group.cards){
      if(!groupsForCard.has(label))groupsForCard.set(label,[]);
      groupsForCard.get(label).push(group.name);
    }
    const badges=cardGroups.map((group,index)=>({attr:'data-card-group="'+index+'"',group,name:group.name,count:group.cards.length}));
    looseCount=masterCards.reduce((count,label)=>count+Number(!groupsForCard.has(label)),0);
    const looseLabel=cardGroups.some(group=>group.name.trim().toLowerCase()==='ungrouped')?'Cards without a group':'Ungrouped';
    badges.push({attr:'data-card-ungrouped',name:looseLabel,count:looseCount});
    cardPicker.set(badges);
  }
  function cardBadge(attribute,name,count){
    return '<button type="button" class="ec-card-badge cursor-interaction" '+attribute+' aria-label="'+esc(name)+'" aria-description="'+count+(count===1?' card':' cards')+'" aria-pressed="false"><span>'+esc(name)+'</span><small class="ec-count" aria-hidden="true">'+count+'</small></button>';
  }
  function syncCardBadgeState(){
    for(const badge of get('ec-card-options').querySelectorAll('[data-card-group]')){
      const group=cardGroups[Number(badge.dataset.cardGroup)],pressed=cardBadgeSelection.has(group);
      badge.setAttribute('aria-pressed',String(pressed));badge.disabled=!pressed&&!group.cards.length;
    }
    const loose=get('ec-card-options').querySelector('[data-card-ungrouped]');
    if(loose){
      loose.setAttribute('aria-pressed',String(usingLooseCards));loose.hidden=!usingLooseCards&&!looseCount;
      const row=loose.closest('.ec-picker-row');if(row&&loose.hidden)row.style.height='0px';
      else if(row)row.style.removeProperty('height');
    }
    get('ec-no-card-groups').hidden=!!masterCards.length||!!cardGroups.length||usingLooseCards;
  }
  function syncCardGroups(list){
    cardBadgeSelection=new Set(selectedCardGroups(list));usingLooseCards=!!list.useUngroupedCards;
    cardPicker.sync();
  }
  function selectedCardGroups(list){
    if(Array.isArray(list.cardGroupNames)){
      if(!list.cardGroupNames.length)return [];
      const byName=new Map(cardGroups.map(group=>[group.name,group]));
      return list.cardGroupNames.map(name=>byName.get(name)).filter(Boolean);
    }
    if(!list.cards.size)return [];
    // Earlier lists saved the cards only. Compact complete groups without changing their selected cards.
    const allSelected=group=>group.cards.length>0&&group.cards.every(label=>list.cards.has(label));
    const covered=new Set(),groups=[];
    for(const group of [...cardGroups].sort((a,b)=>b.cards.length-a.cards.length)){
      if(allSelected(group)&&group.cards.some(label=>!covered.has(label))){
        groups.push(group);group.cards.forEach(label=>covered.add(label));
      }
    }
    return groups;
  }
  function ungroupedCardLabels(){
    const grouped=new Set(cardGroups.flatMap(group=>group.cards));
    return masterCards.filter(label=>!grouped.has(label));
  }
  function additionalCardSelection(list,groups=selectedCardGroups(list),loose=list.useUngroupedCards?ungroupedCardLabels():[]){
    const covered=new Set([...groups.flatMap(group=>group.cards),...loose]);
    return [...list.cards].filter(label=>cardIds.has(label)&&!covered.has(label));
  }
  function cardsFromPools(groups,loose,singles){
    return new Set([...singles,...groups.flatMap(group=>group.cards),...loose].filter(label=>cardIds.has(label)));
  }
  function setCardPoolMembers(list,selected){
    if(list.cards.size===selected.size&&[...selected].every(label=>list.cards.has(label)))return false;
    list.cards=selected;return true;
  }
  function previousCardSelectionChip(list){
    const extras=additionalCardSelection(list);
    return extras.length?chip('card-extras','previous','Previous selection ('+extras.length+')','card-0'):'';
  }
  function changeCardGroupSelection(list,group,selected){
    const groups=selectedCardGroups(list).filter(item=>item!==group);
    if(selected)groups.push(group);
    replaceCardPools(list,groups,list.useUngroupedCards);
  }
  function changeUngroupedSelection(list,selected){
    replaceCardPools(list,selectedCardGroups(list),selected);
  }
  function replaceCardPools(list,groups,useUngroupedCards){
    const singles=additionalCardSelection(list);
    setCardPoolMembers(list,cardsFromPools(groups,useUngroupedCards?ungroupedCardLabels():[],singles));
    list.cardGroupNames=groups.map(group=>group.name);
    list.useUngroupedCards=useUngroupedCards;
  }
  function captureCardRecipes(){
    const recipes=[],loose=ungroupedCardLabels();
    function capture(event,list){
      const groups=selectedCardGroups(list);
      if(!list.cards.size&&!groups.length&&!list.useUngroupedCards)return;
      recipes.push({event,list,groups,singles:additionalCardSelection(list,groups,list.useUngroupedCards?loose:[])});
    }
    // An event or list waiting for Undo should follow the same inventory edits.
    const parents=[...events,...deletedHierarchyItems.filter(entry=>entry.kind==='event').map(entry=>entry.item)];
    for(const event of parents)for(const list of event.lists)capture(event,list);
    for(const entry of deletedHierarchyItems)if(entry.kind==='list'){
      const event=parents.find(parent=>parent.id===entry.parentId);
      if(event)capture(event,entry.item);
    }
    return recipes;
  }
  function refreshCardRecipes(recipes,renamedFrom,renamedTo){
    const loose=ungroupedCardLabels(),currentGroups=new Set(cardGroups);
    for(const {event,list,groups,singles} of recipes){
      // Deleting a named group leaves its existing card assignments available.
      const selected=cardsFromPools(groups,list.useUngroupedCards?loose:[],singles.map(label=>label===renamedFrom?renamedTo:label));
      list.cardGroupNames=groups.filter(group=>currentGroups.has(group)).map(group=>group.name);
      if(setCardPoolMembers(list,selected))rebuildList(event,list);
    }
  }
  drawCardGroups();
  function deleteControl(kind,id,label){
    return '<span class="ec-removal" data-kind="'+esc(kind)+'" data-id="'+esc(id)+'"><button type="button" class="ec-remove-trigger" data-remove-'+kind+'="'+esc(id)+'" aria-label="Remove '+esc(label)+'" aria-expanded="false">×</button><button type="button" class="ec-remove-confirm" aria-label="Confirm removal of '+esc(label)+'" hidden>Confirm</button></span>';
  }
  function chip(type,id,label,color){
    return '<span class="ec-chip '+esc(color)+'"><span class="ec-chip-label">'+esc(label)+'</span>'+deleteControl(type,id,label)+'</span>';
  }
  function addressSelectionChips(list,options){
    const seenGroups=new Set(),chips=[],parents=new Map(addresses.map(parent=>[parent.id,parent]));
    for(const address of options){
      const group=shippingGroupKey(address.parentAddressId);
      if(list.chosen.has(group)){
        if(seenGroups.has(group))continue;
        seenGroups.add(group);
        const parent=parents.get(address.parentAddressId),count=addressPickerCounts.get('g:'+parent.id)[1];
        const label=[parent.person,addressName(parent)].filter(Boolean).join(' · ')+(count>1?' ('+count+')':'');
        chips.push([group,label,'ec-'+parent.color]);
      }else if(list.chosen.has(address.id))chips.push([address.id,[address.person,addressName(address)].filter(Boolean).join(' · '),'ec-'+address.color]);
    }
    return chips;
  }
  function listNeedsRebuild(event,list,options){
    const canFill=requestedProfiles(list)>list.rows.length&&options.some(address=>isDestinationChosen(list,address))
      &&masterCards.some(card=>list.cards.has(card))&&available(event,list).length>0;
    return canFill||disabledDestinationCount(list,allShippingDestinations())>0;
  }
  function renderConfig(rebuild=true,reconcile=true){
    const event=currentEvent(),list=currentList();
    if(!list)return;
    get('ec-pool').checked=list.pool;get('ec-sharing').checked=event.sharing;get('ec-mode').value=list.mode;get('ec-basis').value=list.basis;get('ec-quantity').value=list.quantity;
    const options=activeAddresses();
    updateAddressPickerSelection(list,options);addressChips.set(addressSelectionChips(list,options));
    get('ec-card-chips').innerHTML=previousCardSelectionChip(list);
    syncCardGroups(list);if(rebuild||reconcile&&listNeedsRebuild(event,list,options))refreshLive();else{renderRows();counts();}
  }
  const columnsDialog=get('ec-columns-dialog');
  let columnsOwner=null,fieldCounts=new Map();
  function populatedCounts(rows){
    const counts=new Map();
    for(const row of rows){
      if(Object.entries(billingFields).every(([shipping,billing])=>String(row[shipping]??'')===String(row[billing]??''))){
        const key='sameBillingAndShippingAddress';counts.set(key,(counts.get(key)||0)+1);
      }
      for(const [key] of profileColumns){
        const value=row[key];
        if(value!==false&&String(value??'').trim())counts.set(key,(counts.get(key)||0)+1);
      }
    }
    return counts;
  }
  function omissionSummary(disabled,counts){
    const present=[],payment=[];
    for(const key of disabled){
      const label=fieldLabels.get(key);
      if(paymentColumns.some(([field])=>field===key)){
        if(counts.get('card'))payment.push(label);
      }else if(counts.get(key))present.push(label+' ('+counts.get(key).toLocaleString()+')');
    }
    const notes=[];
    if(present.length)notes.push('Fields containing data will be excluded. '+present.join(', ')+'. Counts show affected profiles.');
    if(payment.length)notes.push('Payment fields left blank for assigned cards include '+payment.join(', ')+'.');
    if(disabled.has('card')&&counts.get('card'))notes.push('Disabling Card also excludes its payment details.');
    return notes.join(' ');
  }
  const uncheckedFields=()=>new Set([...get('ec-field-groups').querySelectorAll('input:not(:checked)')].map(box=>box.value));
  get('ec-open-columns').addEventListener('click',()=>{
    const list=currentList();if(!list)return;
    columnsOwner=list;fieldCounts=populatedCounts(orderedRows(list,true));
    get('ec-field-groups').innerHTML=fieldSections.map(([title,keys])=>'<fieldset><legend>'+title+'</legend>'+keys.map(key=>
      '<label><input type="checkbox" value="'+key+'" '+(list.disabledColumns.has(key)?'':'checked')+'>'+esc(fieldLabels.get(key))+'</label>').join('')+'</fieldset>').join('');
    get('ec-column-impact').textContent=omissionSummary(list.disabledColumns,fieldCounts);
    columnsDialog.showModal();
  });
  get('ec-field-groups').addEventListener('change',()=>{
    get('ec-column-impact').textContent=omissionSummary(uncheckedFields(),fieldCounts);
  });
  get('ec-columns-cancel').addEventListener('click',()=>columnsDialog.close());
  get('ec-columns-apply').addEventListener('click',()=>{
    const list=currentList();
    if(list&&list===columnsOwner){
      list.disabledColumns=uncheckedFields();
      renderRows();saveCoordinatorChange({immediate:true});
    }
    columnsDialog.close();
  });
  columnsDialog.addEventListener('close',()=>{columnsOwner=null;fieldCounts.clear();});
  function setupColumns(list){
    visibleColumns=profileColumns.filter(([key])=>!list.disabledColumns.has(key));
    const heads=get('ec-column-heads'),signature=visibleColumns.map(([key])=>key).join('|');
    if(heads.dataset.layout===signature&&heads.children.length===visibleColumns.length)return;
    heads.dataset.layout=signature;
    get('ec-no-columns').hidden=visibleColumns.length>0;
    for(const button of get('ec-column-jumps').querySelectorAll('[data-jump]')){
      const fields=button.dataset.jump==='nickname'?profileFieldKeys:button.dataset.jump==='firstName'?shippingFieldKeys:billingFieldKeys;
      const target=visibleColumns.find(([key])=>fields.includes(key));
      button.disabled=!target;button.dataset.target=target?.[0]||'';
    }
    get('ec-column-widths').replaceChildren(...visibleColumns.map(()=>document.createElement('col')));
    get('ec-column-heads').innerHTML=visibleColumns.map(([key,label])=>{
      const heading=key==='nickname'?'<button type="button" id="ec-sort-name">'+label+' <span aria-hidden="true" id="ec-sort-glyph"></span></button>':'<span class="ec-column-label">'+label+'</span>';
      const editable=editableFields.has(key);
      const help=[columnHelp[key],editable?'Double-click or press F2 to edit this column for every profile in this list.':''].filter(Boolean).join(' ');
      return '<th scope="col" data-column="'+key+'"'+(help?' title="'+esc(help)+'"':'')+(editable?' tabindex="0" aria-keyshortcuts="F2"':'')+'>'+heading+'<span class="ec-resize-column" role="separator" aria-orientation="vertical" tabindex="0" aria-label="Resize '+label+'" title="Drag to resize · double-click to fit"></span></th>';
    }).join('');
  }
  function applyColumnWidths(){
    const cols=get('ec-column-widths').children,heads=get('ec-column-heads').children,scroller=get('ec-sheet-scroll');
    if(!cols.length)return;
    let total=0,nicknameWidth=0;
    visibleColumns.forEach(([key],index)=>{
      const width=customWidths.get(key)??fittedWidths.get(key)??100;
      cols[index].style.width=width+'px';total+=width;
      heads[index].lastElementChild.setAttribute('aria-valuenow',String(width));
      if(key==='nickname')nicknameWidth=width;
    });
    scroller.style.setProperty('--ec-pin',nicknameWidth+'px');
    scroller.toggleAttribute('data-wide-nickname',!nicknameWidth||nicknameWidth>scroller.clientWidth*0.65);
    get('ec-profile-table').style.width=total+'px';
  }
  function measureFont(element){
    const style=getComputedStyle(element);
    return style.fontStyle+' '+style.fontWeight+' '+style.fontSize+' '+style.fontFamily;
  }
  let columnFitFrame=0;
  function fitSheetColumns(){
    cancelAnimationFrame(columnFitFrame);
    if(columnEditor)return;
    const heads=get('ec-column-heads').children,firstRow=get('ec-rows').querySelector('tr[data-sheet-index]');
    if(!heads.length||!get('ec-sheet-scroll').clientWidth)return;
    const widths=[],fonts=[],measured=new Map();
    for(const [index,[key,label]] of visibleColumns.entries()){
      sheetMeasure.font=measureFont(heads[index]);
      widths[index]=sheetMeasure.measureText(label).width+(key==='nickname'?43:26);
      fonts[index]=firstRow?measureFont(firstRow.cells[index].firstElementChild):sheetMeasure.font;
    }
    function measureRange(col,rows,start,end){
      const key=visibleColumns[col][0];if(flagFields.has(key))return;
      sheetMeasure.font=fonts[col];
      let cache=measured.get(col);
      if(!cache)measured.set(col,cache=new Set());
      for(let i=start;i<end;i++){
        const value=String(sheetCellText(rows[i],key));
        if(cache.has(value))continue;
        cache.add(value);widths[col]=Math.max(widths[col],sheetMeasure.measureText(value).width+26);
      }
    }
    function applyFit(){
      for(const [index,[key]] of visibleColumns.entries())fittedWidths.set(key,Math.ceil(widths[index]));
      applyColumnWidths();
    }
    if(!windowedSheet()){
      for(let col=0;col<widths.length;col++)measureRange(col,sheetRows,0,sheetRows.length);
      applyFit();return;
    }
    // Fit what the user sees immediately, then measure offscreen values in short batches.
    const mounted=[...get('ec-rows').querySelectorAll('tr[data-sheet-index]')].map(tr=>sheetRows[Number(tr.dataset.sheetIndex)]);
    for(let col=0;col<widths.length;col++)measureRange(col,mounted,0,mounted.length);
    applyFit();
    let col=0,offset=0;
    function measureMore(){
      const until=performance.now()+5;
      do{
        const end=Math.min(sheetRows.length,offset+96);
        measureRange(col,sheetRows,offset,end);offset=end;
        if(offset===sheetRows.length){measured.delete(col);col++;offset=0;}
      }while(col<widths.length&&performance.now()<until);
      if(col<widths.length)columnFitFrame=requestAnimationFrame(measureMore);
      else{columnFitFrame=0;applyFit();}
    }
    columnFitFrame=requestAnimationFrame(measureMore);
  }
  let sheetViewportWidth=0;
  new ResizeObserver(()=>{
    const width=get('ec-sheet-scroll').clientWidth;
    if(width&&width!==sheetViewportWidth){sheetViewportWidth=width;fitSheetColumns();}
    scheduleSheetWindow();
  }).observe(get('ec-sheet-scroll'));
  const sheetHeaders=get('ec-column-heads');
  function closeColumnEditor(){
    if(!columnEditor)return;
    const {head,input}=columnEditor;
    columnEditor=null;head.removeAttribute('data-editing');input.remove();
  }
  function openColumnEditor(head){
    const field=head?.dataset.column,list=currentList();
    if(!list||!sheetRows.length||!editableFields.has(field)||columnEditor?.head===head)return;
    closeColumnEditor();
    const flag=flagFields.has(field),input=document.createElement(flag?'select':'input');
    const initial=sheetRows[0][field]??'',uniform=sheetRows.every(row=>(row[field]??'')===initial);
    input.className='ec-column-editor';input.setAttribute('aria-label','Set '+fieldLabels.get(field)+' for all '+sheetRows.length+' profiles');
    if(flag){
      if(!uniform)input.add(new Option('Choose…',''));
      input.add(new Option('Yes','yes'));input.add(new Option('No','no'));
      input.value=uniform?(initial?'yes':'no'):'';
    }else{
      input.type='text';input.autocomplete='off';input.value=uniform?String(initial):'';
      input.placeholder=uniform?'':fieldLabels.get(field);
    }
    // A fit started before this edit can otherwise shift a header near the right edge.
    cancelAnimationFrame(columnFitFrame);
    columnEditor={head,input,field};
    head.setAttribute('data-editing','');head.insertBefore(input,head.lastElementChild);
    function finish(commit){
      if(columnEditor?.input!==input)return;
      const value=input.value;
      closeColumnEditor();
      if(commit&&list===currentList()&&(!flag||value)){
        const targets=sheetRows;
        trackSheetEdit(list,targets.map(row=>row.draftKey??row.accountId),()=>{
          for(const row of targets)applySheetEdits(list,row,{[field]:value});
        });
        renderRows({fitColumns:false});saveCoordinatorChange({immediate:true});
      }
      head.focus({preventScroll:true});
    }
    input.addEventListener('keydown',event=>{
      if(event.isComposing||!['Escape','Enter'].includes(event.key))return;
      event.preventDefault();event.stopPropagation();finish(event.key==='Enter');
    });
    input.addEventListener('blur',()=>{if(columnEditor?.input===input)closeColumnEditor();});
    if(flag)input.addEventListener('change',()=>finish(true));
    input.focus({preventScroll:true});if(!flag)input.select();
  }
  function editableHeader(target){return target.closest('.ec-resize-column,.ec-column-editor')?null:target.closest('th[data-column]');}
  sheetHeaders.addEventListener('dblclick',event=>{const head=editableHeader(event.target);if(head){event.preventDefault();openColumnEditor(head);}});
  // Touch browsers can omit dblclick on table headings.
  let lastHeaderTap=null;
  sheetHeaders.addEventListener('pointerup',event=>{
    if(event.pointerType!=='touch')return;
    const head=editableHeader(event.target),time=event.timeStamp;
    if(head&&lastHeaderTap?.head===head&&time-lastHeaderTap.time<400&&Math.hypot(event.clientX-lastHeaderTap.x,event.clientY-lastHeaderTap.y)<12){
      lastHeaderTap=null;openColumnEditor(head);
    }else lastHeaderTap={head,time,x:event.clientX,y:event.clientY};
  });
  sheetHeaders.addEventListener('keydown',event=>{
    const head=editableHeader(event.target);
    if(!head||!editableFields.has(head.dataset.column)||!(event.key==='F2'||event.key==='Enter'&&event.target===head))return;
    event.preventDefault();openColumnEditor(head);
  });
  sheetHeaders.addEventListener('pointerdown',event=>{
    const handle=event.target.closest('.ec-resize-column');
    if(!handle||event.button!==0)return;
    event.preventDefault();handle.setPointerCapture(event.pointerId);
    const head=handle.parentElement,key=head.dataset.column,start=head.offsetWidth,x=event.clientX;
    function resize(move){
      if(move.pointerId!==event.pointerId)return;
      customWidths.set(key,Math.max(60,Math.round(start+move.clientX-x)));applyColumnWidths();
    }
    function stop(){handle.removeEventListener('pointermove',resize);handle.removeEventListener('lostpointercapture',stop);}
    handle.addEventListener('pointermove',resize);handle.addEventListener('lostpointercapture',stop);
  });
  sheetHeaders.addEventListener('dblclick',event=>{
    const handle=event.target.closest('.ec-resize-column');if(!handle)return;
    customWidths.delete(handle.parentElement.dataset.column);fitSheetColumns();
  });
  sheetHeaders.addEventListener('keydown',event=>{
    const handle=event.target.closest('.ec-resize-column');
    if(!handle||!['ArrowLeft','ArrowRight','Home'].includes(event.key))return;
    event.preventDefault();const head=handle.parentElement,key=head.dataset.column;
    if(event.key==='Home')customWidths.delete(key);
    else customWidths.set(key,Math.max(60,head.offsetWidth+(event.key==='ArrowLeft'?-10:10)));
    applyColumnWidths();
  });
  const sheetRowsByKey=new Map(),sheetRowIndexes=new Map();
  let sheetRows=[],sheetGroups=new Set(),windowFrame=0,sortFrame=0;
  let windowStart=0,windowEnd=0,windowHeight=0,windowRowHeight=35;
  const windowedSheet=()=>sheetRows.length>100;
  const sheetRowKey=row=>row.pending?row.draftKey:String(row.accountId);
  function sheetCellText(row,key){
    if(key==='card'&&row.card)return cardSheetCaption(row.card,sheetGroups);
    if(key==='email'&&!row.email)return row.accountId==null?'Needs account':'Needs email';
    if(key==='card'&&row.pending&&!row.card)return 'Needs card';
    return row[key]??'';
  }
  function rowMarkup(row,index){
    const rowKey=sheetRowKey(row);
    const identity=row.pending?'data-pending aria-label="Pending profile"':'data-account="'+row.accountId+'"';
    return '<tr class="ec-'+esc(row.color)+'" '+identity+' data-sheet-key="'+esc(rowKey)+'" data-sheet-index="'+index+'" aria-rowindex="'+(index+2)+'">'+visibleColumns.map(([key,label])=>{
      const value=esc(sheetCellText(row,key));
      if(readOnlyFields.has(key)){
        const missing=key==='email'&&!row.email||key==='card'&&row.pending&&!row.card;
        return '<td><span class="ec-static-cell'+(missing?' ec-missing-value':'')+'" data-field="'+key+'">'+value+'</span></td>';
      }
      const editing=rowKey?'data-row="'+esc(rowKey)+'"':'readonly tabindex="-1"';
      if(flagFields.has(key))return '<td class="ec-flag-cell"><input type="checkbox" '+editing+' data-field="'+key+'" aria-label="Profile '+(index+1)+' '+label+'" '+(row[key]===true?'checked':'')+'></td>';
      return '<td><input type="text" autocomplete="off" '+editing+' data-field="'+key+'" aria-label="Profile '+(index+1)+' '+label+'" value="'+value+'"></td>';
    }).join('')+'</tr>';
  }
  const compareRows=(list,left,right)=>list.sort*nameCollator.compare(left.nickname,right.nickname)||left.accountId-right.accountId;
  const orderedRows=(list,showPending=false)=>[...list.rows,...(showPending?pendingRows(list):[])].sort((left,right)=>compareRows(list,left,right));
  function indexSheetRows(){
    sheetRowsByKey.clear();sheetRowIndexes.clear();
    sheetRows.forEach((row,index)=>{
      const key=sheetRowKey(row);
      if(key){sheetRowsByKey.set(key,row);sheetRowIndexes.set(key,index);}
    });
  }
  function scheduleSheetWindow(){
    if(windowFrame||!windowedSheet())return;
    windowFrame=requestAnimationFrame(()=>{windowFrame=0;renderSheetWindow();});
  }
  function tableSpacer(count,height,columns){
    const spacer=document.createElement('tr'),cell=spacer.insertCell();
    spacer.className='ec-row-spacer';spacer.setAttribute('aria-hidden','true');
    cell.colSpan=columns;cell.style.height=count*height+'px';
    return spacer;
  }
  function renderSheetWindow(refresh=false,clear=false){
    const body=get('ec-rows'),scroller=get('ec-sheet-scroll'),table=get('ec-profile-table');
    if(!visibleColumns.length||!sheetRows.length){body.replaceChildren();return;}
    const virtual=windowedSheet(),total=sheetRows.length;
    let start=0,end=total;
    if(virtual){
      const height=scroller.clientHeight||window.innerHeight;
      if(!refresh&&windowEnd===total&&height<windowHeight&&Math.abs(scroller.scrollHeight-scroller.scrollTop-windowHeight)<=2){
        scroller.scrollTop=scroller.scrollHeight;
      }
      windowRowHeight=parseFloat(getComputedStyle(table).getPropertyValue('--ec-row-height'))||35;
      const first=Math.min(total-1,Math.max(0,Math.floor(scroller.scrollTop/windowRowHeight)));
      const last=Math.min(total,first+Math.ceil(height/windowRowHeight)+1);
      if(!refresh&&height===windowHeight&&first>=windowStart+(windowStart?3:0)&&last<=windowEnd-(windowEnd===total?0:3))return;
      windowHeight=height;start=Math.max(0,first-9);end=Math.min(total,last+9);
    }
    windowStart=start;windowEnd=end;
    const indices=[];
    for(let i=start;i<end;i++)indices.push(i);
    const active=clear?null:document.activeElement?.closest('#ec-rows tr[data-sheet-index]');
    const pinned=active&&sheetRowIndexes.get(active.dataset.sheetKey);
    if(pinned!=null&&pinned<start)indices.unshift(pinned);
    if(pinned!=null&&pinned>=end)indices.push(pinned);
    const mounted=new Map();
    if(!clear)for(const tr of body.querySelectorAll('tr[data-sheet-index]'))mounted.set(tr.dataset.sheetKey,tr);
    const missing=indices.filter(i=>!mounted.has(sheetRowKey(sheetRows[i])));
    if(missing.length){
      const newRows=document.createElement('tbody');
      newRows.innerHTML=missing.map(i=>rowMarkup(sheetRows[i],i)).join('');
      for(const tr of newRows.rows)mounted.set(tr.dataset.sheetKey,tr);
    }
    const targets=[];let previous=-1;
    for(const i of indices){
      if(i>previous+1)targets.push(tableSpacer(i-previous-1,windowRowHeight,visibleColumns.length));
      const row=mounted.get(sheetRowKey(sheetRows[i]));
      if(Number(row.dataset.sheetIndex)!==i){
        row.dataset.sheetIndex=i;row.setAttribute('aria-rowindex',String(i+2));
        for(const input of row.querySelectorAll('input[data-field]'))input.setAttribute('aria-label','Profile '+(i+1)+' '+fieldLabels.get(input.dataset.field));
      }
      targets.push(row);previous=i;
    }
    if(previous<total-1)targets.push(tableSpacer(total-previous-1,windowRowHeight,visibleColumns.length));
    updateTableRows(body,targets,active);paintSheetSelection();
  }
  function updateTableRows(body,targets,active){
    const keep=new Set(targets);
    for(const row of [...body.rows])if(!keep.has(row))row.remove();
    // Keep a focused input attached, even after a scroll jump or a nickname re-sort.
    const pivot=targets.indexOf(active);
    let before=pivot<0?body.firstChild:active;
    for(let i=pivot-1;i>=0;i--){
      if(targets[i].nextSibling!==before)body.insertBefore(targets[i],before);
      before=targets[i];
    }
    before=pivot<0?body.firstChild:active.nextSibling;
    for(let i=pivot+1;i<targets.length;i++){
      if(targets[i]!==before)body.insertBefore(targets[i],before);
      before=targets[i].nextSibling;
    }
  }
  function sheetInputAt(index,column){
    if(index<0||index>=sheetRows.length)return null;
    const body=get('ec-rows'),scroller=get('ec-sheet-scroll');
    const find=()=>body.querySelector('tr[data-sheet-index="'+index+'"]')?.cells[column]?.querySelector('input[data-row]');
    let input=find();
    if(!input&&windowedSheet()){
      scroller.scrollTop=index*windowRowHeight;
      renderSheetWindow(true);input=find();
    }
    return input;
  }
  function resortSheetRows(){
    if(sortFrame)return;
    sortFrame=requestAnimationFrame(()=>{
      sortFrame=0;const list=currentList();if(!list)return;
      sheetRows.sort((left,right)=>compareRows(list,left,right));indexSheetRows();renderSheetWindow(true);
    });
  }
  get('ec-sheet-scroll').addEventListener('scroll',scheduleSheetWindow,{passive:true});
  function disabledDestinationCount(list,options){
    const disabled=new Set(options.filter(address=>!address.shippingEnabled).map(address=>address.id));
    return list.rows.filter(row=>disabled.has(row.addressId)).length;
  }
  function renderRows({fitColumns=true}={}){
    closeColumnEditor();
    const list=currentList();if(!list)return;setupColumns(list);
    const sortButton=get('ec-sort-name');
    if(sortButton){
      sortButton.closest('th').setAttribute('aria-sort',list.sort===1?'ascending':'descending');
      sortButton.setAttribute('aria-label','Sort profile nickname '+(list.sort===1?'descending':'ascending'));
      get('ec-sort-glyph').textContent=list.sort===1?'↑':'↓';
    }
    sheetRows=orderedRows(list,true);sheetGroups=new Set(list.cardGroupNames||[]);indexSheetRows();
    const table=get('ec-profile-table');
    table.toggleAttribute('data-windowed',windowedSheet());table.setAttribute('aria-rowcount',String(sheetRows.length+1));
    resetSheetSelection();renderSheetWindow(true,true);
    if(fitColumns)fitSheetColumns();
  }
  get('ec-column-heads').addEventListener('click',event=>{if(event.target.closest('#ec-sort-name')){currentList().sort*=-1;renderRows();}});
  get('ec-column-jumps').addEventListener('click',event=>{
    const button=event.target.closest('[data-jump]');if(!button)return;
    const scroller=get('ec-sheet-scroll'),head=get('ec-column-heads'),column=head.querySelector('[data-column="'+button.dataset.target+'"]');
    if(!column)return;
    scroller.scrollLeft=column===head.firstElementChild?0:scroller.scrollLeft+column.getBoundingClientRect().left-scroller.getBoundingClientRect().left-(scroller.hasAttribute('data-wide-nickname')?0:head.firstElementChild.offsetWidth)-1;
  });
  let displayedList=null;
  function renderView(reconcile=true){
    renderEvents();
    const event=currentEvent(),list=currentList(),panel=get('ec-list-panel'),empty=get('ec-hierarchy-empty');
    get('ec-list-strip').hidden=!event;
    panel.hidden=!list;empty.hidden=!!list;
    empty.textContent=event?'Add a list to choose profiles for this event.':'Add an event to get started.';
    if(displayedList!==list){
      displayedList=list;
      get('ec-export-result').textContent='';
      get('ec-status').textContent='';
      get('ec-sheet-scroll').scrollTo(0,0);
    }
    if(!list){
      panel.removeAttribute('aria-labelledby');
      cancelAnimationFrame(columnFitFrame);sheetRows=[];indexSheetRows();
      get('ec-rows').replaceChildren();get('ec-profile-table').setAttribute('aria-rowcount','1');resetSheetSelection();
      return;
    }
    panel.setAttribute('aria-labelledby','ec-tab-'+list.id);
    renderConfig(false,reconcile);
  }
  function allocateAddresses(list,pool,n,random){
    if(!pool.length)return [];
    if(list.mode==='random')return Array.from({length:n},()=>pool[Math.floor(random()*pool.length)]);
    const byOwner=new Map();
    for(const address of pool){
      const key=list.basis==='address'?address.id:address.personId||address.id;
      if(!byOwner.has(key))byOwner.set(key,[]);
      byOwner.get(key).push(address);
    }
    const groups=[...byOwner.values()];
    const cycles=new Map(),result=[];
    while(result.length<n){
      for(const group of shuffle(groups,random)){
        if(!cycles.get(group)?.length)cycles.set(group,shuffle(group,random));
        result.push(cycles.get(group).pop());
        if(result.length===n)break;
      }
    }
    return result;
  }
  function requestedProfiles(list){
    const count=Math.floor(Number(list.quantity)||0);
    return count>0&&count<=100000?count:0;
  }
  function freeAccounts(event,list){
    const eligible=new Map(available(event,list).map(account=>[account.id,account]));
    const returning=[];
    for(const [id] of list.lastAssignments||[]){
      if(eligible.has(id)){returning.push(eligible.get(id));eligible.delete(id);}
    }
    return returning.concat(shuffle(masterAccounts,seededRandom(list.seed)).filter(account=>eligible.has(account.id)));
  }
  function fillDestinations(list,pool,retained,needed){
    if(!needed||!pool.length)return [];
    const held=new Map();
    for(const id of retained)held.set(id,(held.get(id)||0)+1);
    // Count existing assignments toward the distribution before filling open slots.
    const planned=allocateAddresses(list,pool,needed+retained.length,seededRandom(list.seed^0x3593b));
    return planned.filter(address=>{
      const count=held.get(address.id)||0;if(!count)return true;
      held.set(address.id,count-1);return false;
    }).slice(0,needed);
  }
  function pickedCard(list,cards,seed){
    const random=seededRandom(list.seed^Math.imul(seed,0x9e3779b1));
    return cards[Math.floor(random()*cards.length)]||'';
  }
  function pendingRows(list){
    const needed=requestedProfiles(list)-list.rows.length;
    if(needed<=0)return [];
    const pool=activeAddresses().filter(address=>isDestinationChosen(list,address));
    if(!pool.length)return [];
    const eligible=new Set(pool.map(address=>address.id));
    const retained=list.rows.map(row=>row.addressId).filter(id=>eligible.has(id));
    const accounts=freeAccounts(currentEvent(),list),cards=masterCards.filter(card=>list.cards.has(card));
    const counters={...list.nicknameCounters};
    const previous=new Map(list.lastAssignments||[]);
    const owners=accounts.slice(0,needed).map(account=>account.id);
    if(!list.pool&&previous.size){
      const valid=new Set(masterAccounts.filter(account=>accountMatchesPool(account,list)).map(account=>account.id));
      for(const [id,addressId] of previous){
        if(owners.length===needed)break;
        if(valid.has(id)&&eligible.has(addressId))owners.push(id);
      }
    }
    while(owners.length<needed)owners.push(null);
    const plan=fillDestinations(list,pool,retained,needed);
    const counts=new Map();
    for(const row of list.rows)nextAddressNumber(counts,row.addressId);
    return fitRememberedDestinations(list,owners,previous,plan).map((address,index)=>{
      const account=accounts[index],owner=owners[index],key=owner!=null?owner+':'+address.id:null;
      let suffix=key&&list.names.get(key);
      if(!suffix){
        counters[address.id]=(counters[address.id]||0)+1;
        suffix=String(counters[address.id]).padStart(2,'0');
      }
      const card=pickedCard(list,cards,owner??index);
      return {...buildRow(list,account,address,card,nextAddressNumber(counts,address.id),suffix,owner),pending:true};
    });
  }
  function reconcileShippingParent(parent){
    const destinations=shippingDestinationOptions(parent);
    for(const event of events)for(const list of event.lists){
      if(destinations.some(address=>isDestinationChosen(list,address)))rebuildList(event,list);
    }
  }
  function fitRememberedDestinations(list,owners,previous,plan,pinned=new Map()){
    const byId=new Map(plan.map(address=>[address.id,address]));
    const capacity=new Map();
    for(const address of plan)capacity.set(address.id,(capacity.get(address.id)||0)+1);
    // Keep edited rows at their existing addresses first, then keep every other assignment that still fits.
    const keep=owner=>{
      if(owner==null||pinned.has(owner))return;
      const id=previous.get(owner);
      if(!capacity.get(id))return;
      pinned.set(owner,byId.get(id));capacity.set(id,capacity.get(id)-1);
    };
    for(const owner of owners)if(list.edits.has(owner))keep(owner);
    for(const owner of owners)keep(owner);
    const remaining=plan.filter(address=>{
      const count=capacity.get(address.id);if(!count)return false;
      capacity.set(address.id,count-1);return true;
    }).values();
    return owners.map(owner=>pinned.get(owner)||remaining.next().value);
  }
  function rememberAssignments(list){
    if(!list.rows.length)return;
    const recent=new Map(list.rows.map(row=>[row.accountId,row.addressId]));
    for(const [id,address] of list.lastAssignments||[])if(!recent.has(id))recent.set(id,address);
    list.lastAssignments=[...recent];
  }
  function rebuildList(event,list){
    rememberAssignments(list);
    const requested=requestedProfiles(list);
    const cards=masterCards.filter(card=>list.cards.has(card));
    if(!list.pool||!list.chosen.size||!cards.length||!requested){list.rows=[];return;}
    const accountsById=new Map(masterAccounts.map(account=>[account.id,account]));
    const selected=directoryDestinations().filter(address=>isDestinationChosen(list,address));
    const pool=selected.filter(address=>address.active!==false);
    const parents=new Map(directoryAddresses().map(parent=>[parent.id,parent]));
    const inactive=new Map();
    for(const address of selected){
      const parent=parents.get(address.parentAddressId);
      if(parent.active!==false||parent.parcelMissing)continue;
      if(address.shippingEnabled)inactive.set(address.id,address);
    }
    const previous=new Map(list.lastAssignments||[]);
    const usedElsewhere=event.sharing?new Set():otherListAccounts(event,list);
    const retained=list.rows.filter(row=>(pool.length||inactive.has(row.addressId))&&!usedElsewhere.has(row.accountId)).map(row=>accountsById.get(row.accountId)).filter(account=>account&&accountMatchesPool(account,list,true)).slice(0,requested);
    const accounts=pool.length?retained.concat(freeAccounts(event,list).slice(0,requested-retained.length)):retained;
    const owners=accounts.map(account=>account.id),pinned=new Map();
    for(const id of owners){
      const address=inactive.get(previous.get(id));if(address)pinned.set(id,address);
    }
    const plan=allocateAddresses(list,pool,accounts.length-pinned.size,seededRandom(list.seed^0x3593b));
    const assigned=fitRememberedDestinations(list,owners,previous,plan,pinned);
    const counts=new Map();
    list.rows=accounts.map((account,index)=>{
      const address=assigned[index];
      return buildRow(list,account,address,pickedCard(list,cards,account.id),nextAddressNumber(counts,address.id));
    });
  }
  function refreshLive(){const list=currentList();if(!list)return;rebuildList(currentEvent(),list);renderRows();counts();}

  const deletedHierarchyItems=[];
  function clearHierarchyUndo(){
    deletedHierarchyItems.length=0;get('ec-deletion-note').hidden=true;
  }
  function hierarchyChanged(){
    renderView();storeView();saveCoordinatorChange({immediate:true});
  }
  function focusHierarchyTab(kind,id){
    const tab=hierarchyStrip(kind).querySelector('[data-'+kind+'="'+CSS.escape(id||'')+'"]');
    (tab||get(kind==='event'?'ec-new-add':'ec-new-list')).focus();
  }
  function renameHierarchyItem(kind,id){
    const record=hierarchyRecords(kind).find(item=>item.id===id);
    const tab=hierarchyStrip(kind).querySelector('[data-'+kind+'="'+CSS.escape(id)+'"]');
    if(!record||!tab||tab.hidden)return;
    const input=document.createElement('input');input.className='ec-strip-rename';input.type='text';input.maxLength=200;
    input.setAttribute('aria-label','Rename '+kind);input.autocomplete='off';input.value=record.name;
    const editor=document.createElement('span');editor.className='ec-strip-rename-wrap';
    const nameElement=tab.querySelector('.ec-strip-name'),tabWidth=tab.getBoundingClientRect().width;
    const trim=tabWidth-nameElement.getBoundingClientRect().width,maxText=parseFloat(getComputedStyle(nameElement).maxWidth);
    editor.style.width=tabWidth+'px';editor.style.font=getComputedStyle(tab).font;
    editor.append(input,tab.querySelector('.ec-count').cloneNode(true));tab.before(editor);tab.hidden=true;
    input.addEventListener('input',()=>{
      sheetMeasure.font=measureFont(input);
      editor.style.width=Math.ceil(trim+Math.min(Number.isFinite(maxText)?maxText:320,sheetMeasure.measureText(input.value).width))+'px';
    });
    let closed=false;
    function finish(cancel,focus){
      if(closed)return;
      closed=true;editor.remove();tab.hidden=false;
      const name=input.value.trim();
      if(!cancel&&name&&name!==record.name){record.name=name;hierarchyChanged();}
      if(focus)tab.focus();
    }
    input.addEventListener('blur',()=>finish(false,false));
    input.addEventListener('keydown',event=>{
      if(event.isComposing||!['Enter','Escape'].includes(event.key))return;
      event.preventDefault();event.stopPropagation();finish(event.key==='Escape',true);
    });
    input.focus();input.select();
  }
  function addHierarchyItem(kind){
    const records=hierarchyRecords(kind),event=currentEvent(),stem=kind==='event'?'Event':'List';
    if(kind==='list'&&!event)return;
    const names=new Set(records.map(item=>item.name.toLowerCase()));
    let n=1;while(names.has((stem+' '+n).toLowerCase()))n++;
    const name=stem+' '+n;
    const item=kind==='event'?makeEvent('event-'+eventSerial++,name):makeList(nextListId(event),name,'sniped');
    records.push(item);
    if(kind==='event')activeEvent=item.id;else event.selectedListId=item.id;
    hierarchyChanged();renameHierarchyItem(kind,item.id);
  }
  function selectHierarchyItem(kind,id){
    if(!hierarchyRecords(kind).some(record=>record.id===id))return;
    if(kind==='event')activeEvent=id;else currentEvent().selectedListId=id;
    get('ec-status').textContent='';renderView(false);storeView();
  }
  function showHierarchyUndo(){
    const latest=deletedHierarchyItems.at(-1);get('ec-deletion-note').hidden=!latest;
    get('ec-deletion-text').textContent=latest?'Deleted '+latest.item.name+'.':'';
  }
  function removeHierarchyItem(kind,id){
    const records=hierarchyRecords(kind),index=records.findIndex(record=>record.id===id);
    if(index<0)return;
    const parent=currentEvent(),[item]=records.splice(index,1);
    deletedHierarchyItems.push({kind,item,index,parentId:kind==='list'?parent.id:null});
    const fallback=records[index]?.id||records[index-1]?.id||null;
    if(kind==='event'&&activeEvent===id)activeEvent=fallback;
    if(kind==='list'&&parent.selectedListId===id)parent.selectedListId=fallback;
    hierarchyChanged();showHierarchyUndo();get('ec-undo-delete').focus();
  }
  get('ec-undo-delete').addEventListener('click',()=>{
    const deleted=deletedHierarchyItems.at(-1);if(!deleted)return;
    const {kind,item,index,parentId}=deleted,parent=kind==='list'?events.find(event=>event.id===parentId):null;
    const records=kind==='event'?events:parent?.lists;if(!records)return;
    deletedHierarchyItems.pop();records.splice(Math.min(index,records.length),0,item);
    if(parent){activeEvent=parent.id;parent.selectedListId=item.id;}else activeEvent=item.id;
    hierarchyChanged();showHierarchyUndo();focusHierarchyTab(kind,item.id);
  });
  for(const kind of ['event','list']){
    const strip=hierarchyStrip(kind);
    get(kind==='event'?'ec-new-add':'ec-new-list').addEventListener('click',()=>addHierarchyItem(kind));
    strip.addEventListener('click',event=>{
      const tab=event.target.closest('[role=tab]');
      if(tab&&tab.getAttribute('aria-selected')!=='true')selectHierarchyItem(kind,tab.dataset[kind]);
    });
    strip.addEventListener('dblclick',event=>{
      const tab=event.target.closest('[role=tab]');if(tab)renameHierarchyItem(kind,tab.dataset[kind]);
    });
    strip.addEventListener('keydown',event=>{
      const tab=event.target.closest('[role=tab]');if(!tab)return;
      if(event.key==='F2'){event.preventDefault();renameHierarchyItem(kind,tab.dataset[kind]);return;}
      const tabs=[...strip.querySelectorAll('[role=tab]')],i=tabs.indexOf(tab);
      const at={Home:0,End:tabs.length-1,ArrowLeft:(i+tabs.length-1)%tabs.length,ArrowRight:(i+1)%tabs.length}[event.key];
      if(at===undefined)return;
      event.preventDefault();const id=tabs[at].dataset[kind];selectHierarchyItem(kind,id);focusHierarchyTab(kind,id);
    });
    strip.addEventListener('ec:confirmed-remove',event=>{
      if(event.detail?.kind===kind)removeHierarchyItem(kind,event.detail.id);
    });
  }
  get('ec-address-options').addEventListener('change',e=>{
    const list=currentList(),{parent,destinationGroup,address:id}=e.target.dataset,checked=e.target.checked;
    const destinations=directoryDestinations();
    if(parent||destinationGroup){
      const matching=destinations.filter(a=>parent?(a.personId||'__unassigned__')===parent:a.parentAddressId===destinationGroup);
      for(const address of matching){
        list.chosen.delete(address.id);
        const group=shippingGroupKey(address.parentAddressId);
        if(!checked)list.chosen.delete(group);
        else if(address.active!==false)list.chosen.add(group);
      }
    }else if(id){
      const address=destinations.find(option=>option.id===id);if(!address)return;
      if(list.chosen.delete(shippingGroupKey(address.parentAddressId))){
        const inUse=new Set(list.rows.map(row=>row.addressId));
        for(const sibling of destinations)if(sibling.parentAddressId===address.parentAddressId&&(sibling.active!==false||inUse.has(sibling.id)))list.chosen.add(sibling.id);
      }
      if(checked)list.chosen.add(id);else list.chosen.delete(id);
    }
    renderConfig();get('ec-status').textContent='';
  });
  get('ec-card-options').addEventListener('click',event=>{
    const badge=event.target.closest('.ec-card-badge'),list=currentList();
    if(!badge||!list)return;
    const selected=badge.getAttribute('aria-pressed')!=='true';
    if(badge.hasAttribute('data-card-ungrouped'))changeUngroupedSelection(list,selected);
    else changeCardGroupSelection(list,cardGroups[Number(badge.dataset.cardGroup)],selected);
    renderConfig();get('ec-status').textContent='';
  });
  const removalWindowMs=2000;
  let armedChip=null,armDeadline=0,armTimer;
  function disarm(){
    clearTimeout(armTimer);if(!armedChip)return;
    const trigger=armedChip.querySelector('.ec-remove-trigger'),confirm=armedChip.querySelector('.ec-remove-confirm');
    if(document.activeElement===confirm)trigger.focus();
    trigger.removeAttribute('data-armed');trigger.setAttribute('aria-expanded','false');
    confirm.hidden=true;armedChip=null;
  }
  root.addEventListener('click',event=>{
    const button=event.target.closest('.ec-removal button');if(!button)return;
    event.stopPropagation();const control=button.closest('.ec-removal');
    if(button.classList.contains('ec-remove-trigger')){
      if(control===armedChip)return;
      disarm();armedChip=control;armDeadline=performance.now()+removalWindowMs;
      button.dataset.armed='';button.setAttribute('aria-expanded','true');
      control.querySelector('.ec-remove-confirm').hidden=false;armTimer=setTimeout(disarm,removalWindowMs);return;
    }
    if(control!==armedChip||performance.now()>=armDeadline){disarm();return;}
    const {kind,id}=control.dataset;disarm();
    control.querySelector('.ec-remove-trigger').dispatchEvent(new CustomEvent('ec:confirmed-remove',{bubbles:true,detail:{kind,id}}));
  },true);
  let pendingDirectoryRemoval=null;
  function closeDirectoryRemoval(panel){
    const pending=pendingDirectoryRemoval;
    if(!pending||panel&&panel!==pending.panel)return;
    clearTimeout(pending.timer);pendingDirectoryRemoval=null;
    if(pending.panel.contains(document.activeElement))pending.trigger.focus();
    pending.panel.hidden=true;
  }
  function openDirectoryRemoval(trigger,panel,cancel){
    closeDirectoryRemoval();
    panel.hidden=false;
    pendingDirectoryRemoval={trigger,panel,deadline:performance.now()+removalWindowMs,timer:setTimeout(closeDirectoryRemoval,removalWindowMs)};
    cancel.focus();
  }
  function confirmDirectoryRemoval(panel){
    const active=pendingDirectoryRemoval?.panel===panel&&!panel.hidden&&performance.now()<pendingDirectoryRemoval.deadline;
    closeDirectoryRemoval(panel);
    return active;
  }
  function removeChip(event){
    const button=event.target.closest('[data-remove-address],[data-remove-card-extras],[data-remove-recipient],[data-remove-pool]');if(!button)return;
    if(button.hasAttribute('data-remove-card-extras')){
      const list=currentList();
      for(const label of additionalCardSelection(list))list.cards.delete(label);
      renderConfig();return;
    }
    if(button.hasAttribute('data-remove-pool')){currentList().pool=false;renderConfig();return;}
    if(button.hasAttribute('data-remove-recipient')){
      const form=button.form;
      const input=form.elements.namedItem('newRecipient');
      if(input.dataset.editingRecipient===button.dataset.removeRecipient)clearRecipientInput(input);
      button.closest('.ec-chip').remove();
      saveRecipientChips(form);
      input.focus();
      return;
    }
    currentList().chosen.delete(button.dataset.removeAddress);renderConfig();
  }
  for(const id of ['ec-address-chips','ec-account-chips','ec-card-chips','ec-source-editor'])get(id).addEventListener('ec:confirmed-remove',removeChip);
  const pickers=[...root.querySelectorAll('details')];
  function closeOthers(target){for(const picker of pickers)if(picker.open&&!picker.contains(target))picker.open=false;}
  document.addEventListener('pointerdown',event=>{closeOthers(event.target);if(armedChip&&!armedChip.contains(event.target))disarm();});
  root.addEventListener('focusin',event=>{closeOthers(event.target);if(armedChip&&!armedChip.contains(event.target))disarm();});
  root.addEventListener('keydown',event=>{if(event.key==='Escape'){disarm();closeDirectoryRemoval();const picker=event.target.closest('details[open]');if(picker){picker.open=false;picker.querySelector('summary').focus();}}});
  get('ec-pool').addEventListener('change',()=>{currentList().pool=get('ec-pool').checked;refreshLive();});
  get('ec-sharing').addEventListener('change',()=>{currentEvent().sharing=get('ec-sharing').checked;refreshLive();});
  get('ec-mode').addEventListener('change',()=>{currentList().mode=get('ec-mode').value;refreshLive();});get('ec-basis').addEventListener('change',()=>{currentList().basis=get('ec-basis').value;refreshLive();});
  get('ec-quantity').addEventListener('input',()=>{currentList().quantity=get('ec-quantity').value;refreshLive();});
  get('ec-select-unused').addEventListener('click',()=>{const list=currentList();list.quantity=list.rows.length+available().length;get('ec-quantity').value=list.quantity;refreshLive();});
  const sheetBody=get('ec-rows'),bulkValue=get('ec-cell-value');
  let selectedCells=new Map(),selectedCellCount=0;
  const sheetScroller=get('ec-sheet-scroll');
  sheetScroller.tabIndex=-1;
  const compactSheet=matchMedia('(max-width:1024px),(any-pointer:coarse) and (max-width:1390px)');
  const setup=get('ec-list-setup'),setupToggle=get('ec-toggle-setup');
  setupToggle.addEventListener('click',()=>{
    setup.hidden=!setup.hidden;
    if(setup.hidden)for(const picker of setup.querySelectorAll('details[open]'))picker.open=false;
    setupToggle.setAttribute('aria-expanded',String(!setup.hidden));
    setupToggle.textContent=setup.hidden?'Show setup':'Hide setup';
    scheduleSheetHeight();
  });
  function positionSetupPicker(){
    if(!compactSheet.matches)return;
    const picker=setup.querySelector('details[open]');if(!picker)return;
    const anchor=picker.querySelector('summary').getBoundingClientRect(),port=setup.getBoundingClientRect();
    if(anchor.bottom<=port.top||anchor.top>=port.bottom){picker.open=false;return;}
    const vp=window.visualViewport,top=vp?.offsetTop??0,left=vp?.offsetLeft??0;
    const bottom=top+(vp?.height??innerHeight),width=vp?.width??innerWidth;
    const below=bottom-anchor.bottom-8,above=anchor.top-top-8,opensDown=below>=Math.min(220,above);
    const menu=picker.querySelector('.ec-options'),size=Math.min(anchor.width,width-16);
    menu.style.setProperty('--ec-menu-width',size+'px');
    menu.style.setProperty('--ec-menu-left',Math.max(left+8,Math.min(anchor.left,left+width-8-size))+'px');
    menu.style.setProperty('--ec-menu-top',opensDown?anchor.bottom+4+'px':'auto');
    menu.style.setProperty('--ec-menu-bottom',opensDown?'auto':innerHeight-anchor.top+4+'px');
    menu.style.setProperty('--ec-menu-height',Math.min(352,opensDown?below:above)+'px');
  }
  setup.addEventListener('toggle',positionSetupPicker,true);
  setup.addEventListener('scroll',positionSetupPicker,{passive:true});
  let heightFrame=0;
  function scheduleSheetHeight(){
    if(heightFrame)return;
    heightFrame=requestAnimationFrame(()=>{
      heightFrame=0;
      const viewport=window.visualViewport;
      const height=viewport?.height??window.innerHeight;
      if(compactSheet.matches){
        root.style.setProperty('--ec-visual-height',height+'px');
        root.style.setProperty('--ec-visual-top',(viewport?.offsetTop??0)+'px');
        positionSetupPicker();return;
      }
      const rect=sheetScroller.getBoundingClientRect();
      if(!rect.width)return;
      const top=Math.max(0,rect.top-(viewport?.offsetTop??0));
      const footer=root.getBoundingClientRect().bottom-rect.bottom;
      // On a stacked/mobile layout the sheet can start below the fold. Keep enough
      // rows to scroll while the user brings it into view.
      const available=Math.max(Math.min(180,height/3),Math.floor(height-top-footer));
      sheetScroller.style.setProperty('--ec-sheet-height',available+'px');
    });
  }
  new ResizeObserver(scheduleSheetHeight).observe(root);
  window.addEventListener('resize',scheduleSheetHeight);
  window.addEventListener('scroll',scheduleSheetHeight,{passive:true});
  window.visualViewport?.addEventListener('resize',scheduleSheetHeight);
  window.visualViewport?.addEventListener('scroll',scheduleSheetHeight);
  let cellAnchor=null,selectingCells=false;
  let cellDrag=null,dragFrame=0;
  const formatButton=get('ec-format-address');
  let formatButtonEligible=false,formatButtonFrame=0,formatPoint=null;
  function visibleFormatCell(input,bounds){
    if(!input?.classList.contains('ec-selected-cell')||!profileFormatFields.has(input.dataset.field))return null;
    const rect=input.getBoundingClientRect();
    const left=Math.max(rect.left,bounds.left),right=Math.min(rect.right,bounds.right);
    const top=Math.max(rect.top,bounds.top),bottom=Math.min(rect.bottom,bounds.bottom);
    return right-left>=8&&bottom-top>=8?{left,right,top,bottom}:null;
  }
  function positionFormatButton(){
    if(!formatButtonEligible||cellDrag?.active){formatButton.hidden=true;return;}
    const viewport=window.visualViewport,viewportLeft=viewport?.offsetLeft??0,viewportTop=viewport?.offsetTop??0;
    const box=sheetScroller.getBoundingClientRect();
    const pinned=sheetScroller.hasAttribute('data-wide-nickname')?box.left:sheetHeaders.firstElementChild?.getBoundingClientRect().right??box.left;
    const bounds={
      left:Math.max(viewportLeft,box.left+sheetScroller.clientLeft,pinned),
      right:Math.min(viewportLeft+(viewport?.width??innerWidth),box.left+sheetScroller.clientLeft+sheetScroller.clientWidth),
      top:Math.max(viewportTop,box.top+sheetScroller.clientTop,sheetHeaders.getBoundingClientRect().bottom),
      bottom:Math.min(viewportTop+(viewport?.height??innerHeight),box.top+sheetScroller.clientTop+sheetScroller.clientHeight)
    };
    const row=formatPoint&&sheetBody.querySelector('tr[data-sheet-key="'+CSS.escape(formatPoint.row)+'"]');
    let anchor=visibleFormatCell(row?.cells[formatPoint?.column]?.querySelector('input'),bounds);
    if(!anchor){
      const cells=sheetBody.querySelectorAll('input.ec-selected-cell');
      for(let i=cells.length-1;i>=0&&!anchor;i--)anchor=visibleFormatCell(cells[i],bounds);
    }
    if(!anchor){formatButton.hidden=true;return;}
    formatButton.hidden=false;
    const width=formatButton.offsetWidth,height=formatButton.offsetHeight,gap=6;
    if(bounds.right-bounds.left<width||bounds.bottom-bounds.top<height){formatButton.hidden=true;return;}
    const clamp=(value,start,end)=>Math.max(start,Math.min(value,end));
    const maxX=bounds.right-width,maxY=bounds.bottom-height;
    let x=clamp(anchor.left,bounds.left,maxX);
    let y=clamp((anchor.top+anchor.bottom-height)/2,bounds.top,maxY);
    if(anchor.right+gap<=maxX)x=anchor.right+gap;
    else if(anchor.left-gap-width>=bounds.left)x=anchor.left-gap-width;
    else if(anchor.bottom+gap<=maxY)y=anchor.bottom+gap;
    else y=clamp(anchor.top-gap-height,bounds.top,maxY);
    formatButton.style.left=x+'px';
    formatButton.style.top=y+'px';
  }
  function scheduleFormatButton(){
    if(formatButtonFrame)return;
    formatButtonFrame=requestAnimationFrame(()=>{formatButtonFrame=0;positionFormatButton();});
  }
  sheetScroller.addEventListener('scroll',scheduleFormatButton,{passive:true});
  window.addEventListener('scroll',scheduleFormatButton,{passive:true});
  window.addEventListener('resize',scheduleFormatButton);
  window.visualViewport?.addEventListener('scroll',scheduleFormatButton,{passive:true});
  window.visualViewport?.addEventListener('resize',scheduleFormatButton);
  new ResizeObserver(scheduleFormatButton).observe(sheetScroller);
  function showSelection(){
    const count=selectedCellCount,editing=!bulkValue.hidden;
    updateSheetHistoryControls();
    sheetScroller.toggleAttribute('data-picking-cells',selectingCells);
    get('ec-select-cells').setAttribute('aria-pressed',String(selectingCells));
    get('ec-match-billing').hidden=!count||editing;
    formatButtonEligible=!editing&&count>0&&[...selectedCells.values()].some(fields=>[...fields].some(field=>profileFormatFields.has(field)));
    if(!formatButtonEligible)formatButton.hidden=true;
    scheduleFormatButton();
    get('ec-edit-cells').hidden=!count||!bulkValue.hidden||(!selectingCells&&count<2);
    get('ec-apply-cells').hidden=!editing;get('ec-cancel-cells').hidden=!editing;
    get('ec-selection-info').textContent=count?count+(count===1?' cell':' cells')+' selected':selectingCells?'Tap cells to select them.':'';
  }
  function paintSheetSelection(){
    for(const tr of sheetBody.rows){
      if(!tr.hasAttribute('data-sheet-index'))continue;
      const row=sheetRows[Number(tr.dataset.sheetIndex)];
      const fields=row&&selectedCells.get(sheetRowKey(row));
      if(!fields&&!tr.hasAttribute('data-has-selection'))continue;
      tr.toggleAttribute('data-has-selection',!!fields);
      for(const td of tr.cells){
        const input=td.firstElementChild;
        if(input?.hasAttribute('data-row'))input.classList.toggle('ec-selected-cell',!!fields?.has(input.dataset.field));
      }
    }
    scheduleFormatButton();
  }
  function setSheetSelection(cells,count){
    selectedCells=cells;selectedCellCount=count;
    paintSheetSelection();bulkValue.hidden=true;showSelection();
  }
  function resetSheetSelection(){
    stopCellDrag();cellAnchor=null;formatPoint=null;selectingCells=false;setSheetSelection(new Map(),0);
  }
  function sheetPoint(td){
    const index=td?.parentElement?.dataset.sheetIndex;
    if(index===undefined||!sheetRows[index])return null;
    return {row:sheetRowKey(sheetRows[index]),column:td.cellIndex};
  }
  function selectSingleCell(input){
    const {row,field}=input.dataset;
    cellAnchor=sheetPoint(input.closest('td'));formatPoint=cellAnchor;
    setSheetSelection(new Map([[row,new Set([field])]]),1);
  }
  function toggleSelectedCell(input){
    const {row,field}=input.dataset,fields=new Set(selectedCells.get(row));
    const adding=!fields.delete(field);
    if(adding)fields.add(field);
    if(fields.size)selectedCells.set(row,fields);else selectedCells.delete(row);
    cellAnchor=sheetPoint(input.closest('td'));formatPoint=cellAnchor;
    setSheetSelection(selectedCells,selectedCellCount+(adding?1:-1));
  }
  function selectRange(from,to){
    if(!from||!to)return;
    const a=sheetRowIndexes.get(from.row),b=sheetRowIndexes.get(to.row);
    if(a===undefined||b===undefined)return;
    const firstCol=Math.min(from.column,to.column),lastCol=Math.max(from.column,to.column);
    const fields=new Set(visibleColumns.slice(firstCol,lastCol+1).map(([field])=>field).filter(field=>editableFields.has(field)));
    const cells=new Map();
    if(fields.size){
      for(let i=Math.min(a,b);i<=Math.max(a,b);i++){
        const row=sheetRowKey(sheetRows[i]);
        if(row)cells.set(row,fields);
      }
    }
    formatPoint=to;
    setSheetSelection(cells,cells.size*fields.size);
  }
  function stopCellDrag(){
    if(!cellDrag)return;
    const id=cellDrag.id;cellDrag=null;
    cancelAnimationFrame(dragFrame);dragFrame=0;
    sheetScroller.removeAttribute('data-dragging-cells');
    if(sheetScroller.hasPointerCapture(id))sheetScroller.releasePointerCapture(id);
    scheduleFormatButton();
  }
  function edgeStep(point,start,end){
    const band=Math.min(28,(end-start)/4);
    if(point<start+band)return -Math.min(22,Math.ceil((start+band-point)/2));
    if(point>end-band)return Math.min(22,Math.ceil((point-end+band)/2));
    return 0;
  }
  function paintDrag(){
    dragFrame=0;
    const d=cellDrag;if(!d?.active)return;
    const box=sheetScroller.getBoundingClientRect();
    const left=Math.max(0,box.left+sheetScroller.clientLeft),right=Math.min(innerWidth,box.left+sheetScroller.clientLeft+sheetScroller.clientWidth);
    const top=Math.max(0,sheetHeaders.getBoundingClientRect().bottom),bottom=Math.min(innerHeight,box.top+sheetScroller.clientTop+sheetScroller.clientHeight);
    if(bottom<=top||right<=left)return;
    const dx=edgeStep(d.x,left,right),dy=edgeStep(d.y,top,bottom);
    const beforeX=sheetScroller.scrollLeft,beforeY=sheetScroller.scrollTop;
    if(dx||dy)sheetScroller.scrollBy(dx,dy);
    if(sheetScroller.scrollTop!==beforeY)renderSheetWindow();
    const x=Math.max(left+1,Math.min(right-1,d.x)),y=Math.max(top+1,Math.min(bottom-1,d.y));
    const cell=document.elementFromPoint(x,y)?.closest('td');
    const end=cell&&sheetBody.contains(cell)?sheetPoint(cell):null;
    if(end&&(end.row!==d.end?.row||end.column!==d.end?.column)){
      d.end=end;selectRange(d.start,end);
    }
    if(sheetScroller.scrollLeft!==beforeX||sheetScroller.scrollTop!==beforeY)dragFrame=requestAnimationFrame(paintDrag);
  }
  sheetBody.addEventListener('pointerdown',event=>{
    const cell=event.target.closest('td');if(!cell||event.button!==0||!event.isPrimary)return;
    const point=sheetPoint(cell);if(!point)return;
    const input=cell.querySelector('input[data-row]');
    if(!event.shiftKey&&!event.ctrlKey&&!event.metaKey&&(event.pointerType!=='touch'||selectingCells)){
      stopCellDrag();cellDrag={id:event.pointerId,start:point,x:event.clientX,y:event.clientY,active:false};
    }
    if(!input)return;
    if(event.shiftKey&&cellAnchor){event.preventDefault();selectRange(cellAnchor,point);}
    else if(selectingCells||event.ctrlKey||event.metaKey){event.preventDefault();toggleSelectedCell(input);}
    else{selectSingleCell(input);return;}
    sheetScroller.focus({preventScroll:true});
  });
  document.addEventListener('pointermove',event=>{
    const d=cellDrag;if(!d||d.id!==event.pointerId)return;
    if(!(event.buttons&1)){stopCellDrag();return;}
    if(!d.active){
      if(Math.hypot(event.clientX-d.x,event.clientY-d.y)<5)return;
      const target=document.elementFromPoint(event.clientX,event.clientY)?.closest('td');
      const point=sheetBody.contains(target)?sheetPoint(target):null;
      if(point?.row===d.start.row&&point.column===d.start.column)return; // Preserve normal text selection inside one input.
      d.active=true;cellAnchor=d.start;formatButton.hidden=true;
      sheetScroller.setPointerCapture(d.id);sheetScroller.setAttribute('data-dragging-cells','');
      sheetScroller.focus({preventScroll:true});
    }
    event.preventDefault();d.x=event.clientX;d.y=event.clientY;
    if(!dragFrame)dragFrame=requestAnimationFrame(paintDrag);
  });
  document.addEventListener('pointerup',event=>{
    if(cellDrag?.id!==event.pointerId)return;
    if(cellDrag.active){
      cellDrag.x=event.clientX;cellDrag.y=event.clientY;
      cancelAnimationFrame(dragFrame);paintDrag();
    }
    stopCellDrag();
  });
  document.addEventListener('pointercancel',event=>{if(cellDrag?.id===event.pointerId)stopCellDrag();});
  sheetScroller.addEventListener('lostpointercapture',event=>{
    if(event.target===sheetScroller&&cellDrag?.id===event.pointerId)stopCellDrag();
  });
  sheetBody.addEventListener('dragstart',event=>{if(cellDrag)event.preventDefault();});
  window.addEventListener('blur',stopCellDrag);
  sheetBody.addEventListener('focusin',event=>{
    const input=event.target.closest('input[data-row]');
    if(input&&!selectedCells.get(input.dataset.row)?.has(input.dataset.field))selectSingleCell(input);
  });
  get('ec-select-cells').addEventListener('click',()=>{
    const enable=!selectingCells;
    resetSheetSelection();selectingCells=enable;showSelection();
    if(enable)sheetScroller.focus({preventScroll:true});
  });
  function startBulk(value){
    if(!selectedCellCount)return;
    bulkValue.hidden=false;bulkValue.value=value??'';showSelection();
    bulkValue.focus({preventScroll:true});
  }
  get('ec-edit-cells').addEventListener('click',()=>startBulk());
  function updateSheetHistoryControls(){
    const list=currentList(),history=list&&sheetHistories.get(list);
    get('ec-sheet-undo').disabled=!history?.undo.length;
    get('ec-sheet-redo').disabled=!history?.redo.length;
    get('ec-sheet-defaults').disabled=!list?.edits.size;
  }
  function redrawSheetEdits(list){
    typingAction=null;refreshSourceRows(null,list);renderRows();
    saveCoordinatorChange({immediate:true});sheetScroller.focus({preventScroll:true});
  }
  function replaySheetEdit(direction){
    const list=currentList(),history=list&&sheetHistories.get(list);
    const action=history?.[direction].pop();if(!action)return;
    for(const [key,change] of action)setSheetOverrides(list,key,overrideStep(list.edits.get(key),change,direction));
    history[direction==='undo'?'redo':'undo'].push(action);
    redrawSheetEdits(list);
  }
  get('ec-sheet-undo').addEventListener('click',()=>replaySheetEdit('undo'));
  get('ec-sheet-redo').addEventListener('click',()=>replaySheetEdit('redo'));
  const defaultsDialog=get('ec-defaults-dialog'),defaultsScope=get('ec-defaults-scope');
  let restoreList=null,restoreSelection=new Map();
  get('ec-sheet-defaults').addEventListener('click',()=>{
    restoreList=currentList();if(!restoreList)return;
    restoreSelection=new Map();
    for(const [key,fields] of selectedCells){
      const row=sheetRowsByKey.get(key);
      if(row)restoreSelection.set(row,fields);
    }
    const cells=selectedCellCount,rows=restoreSelection.size;
    defaultsScope.replaceChildren();
    if(cells){
      defaultsScope.add(new Option(cells===1?'Selected cell':cells+' selected cells','cells'));
      defaultsScope.add(new Option(rows===1?'Entire selected profile':rows+' entire selected profiles','profiles'));
    }
    defaultsScope.add(new Option('All profiles in this list','list'));
    defaultsDialog.showModal();
  });
  get('ec-defaults-cancel').addEventListener('click',()=>defaultsDialog.close());
  defaultsDialog.addEventListener('close',()=>{restoreList=null;restoreSelection.clear();});
  function removeSheetOverrides(list,key,fields){
    if(!fields){list.edits.delete(key);return;}
    const saved=list.edits.get(key);if(!saved)return;
    const edits={...saved},billing=Object.values(billingFields);
    const defaults=new Set(edits.defaultBillingFields||[]);
    for(const field of fields){
      delete edits[field];
      if(edits.matchShippingBilling&&billing.includes(field))defaults.add(field);
    }
    if(defaults.size===billing.length){delete edits.matchShippingBilling;delete edits.defaultBillingFields;}
    else if(defaults.size)edits.defaultBillingFields=[...defaults];
    setSheetOverrides(list,key,edits);
  }
  get('ec-defaults-apply').addEventListener('click',()=>{
    const list=currentList(),scope=defaultsScope.value;
    if(list&&list===restoreList){
      const targets=new Map();
      if(scope==='list')for(const key of list.edits.keys())targets.set(key,null);
      else for(const [row,fields] of restoreSelection){
        for(const key of [row.sourceAccountId??row.accountId,row.draftKey]){
          if(key==null)continue;
          if(scope==='profiles')targets.set(key,null);
          else targets.set(key,new Set([...(targets.get(key)||[]),...fields]));
        }
      }
      trackSheetEdit(list,targets.keys(),()=>{
        for(const [key,fields] of targets)removeSheetOverrides(list,key,fields);
      });
      redrawSheetEdits(list);
    }
    defaultsDialog.close();
  });

  const shippingStreetFields=['street','street2'];
  const profileFormatGroups={
    street:[...shippingStreetFields,'billingStreet','billingStreet2'],
    name:['firstName','lastName','billingFirstName','billingLastName'],
    phone:['phone','billingPhone']
  };
  const profileFormatFields=new Map(Object.entries(profileFormatGroups).flatMap(([group,fields])=>fields.map(field=>[field,group])));
  const streetLineOne=['street','billingStreet'];
  const profileFormatOptions={
    direction:streetLineOne,suffix:streetLineOne,unit:profileFormatGroups.street,
    name:profileFormatGroups.name,phone:profileFormatGroups.phone
  };
  const formatDialog=get('ec-format-dialog'),formatPreview=get('ec-format-preview');
  let formatOwner=null,formatSelection=new Map();
  function postalWords(words){
    const spellings=new Map();
    for(const [long,short] of Object.entries(words)){
      spellings.set(long.toLowerCase(),short);
      spellings.set(short.toLowerCase(),short);
    }
    return spellings;
  }
  const postalDirections=postalWords({North:'N',South:'S',East:'E',West:'W',Northeast:'NE',Northwest:'NW',Southeast:'SE',Southwest:'SW'});
  const postalSuffixes=postalWords({Street:'St',Avenue:'Ave',Boulevard:'Blvd',Circle:'Cir',Court:'Ct',Drive:'Dr',Highway:'Hwy',Lane:'Ln',Parkway:'Pkwy',Place:'Pl',Road:'Rd',Terrace:'Ter',Trail:'Trl',Way:'Way'});
  const postalUnits=postalWords({Apartment:'Apt',Suite:'Ste'});
  const postalLookup=(words,value)=>words.get(value.toLowerCase().replace(/\./g,''));
  const postalSpelling=(words,value,enabled)=>enabled?(postalLookup(words,value)??value):value;
  function formatPostalStreet(value,field,modes){
    const text=value.trimEnd(),trailing=value.slice(text.length),unit=text.match(unitAtEnd);
    const unitText=unit?text.slice(unit.index):'';
    const formattedUnit=unitText.replace(/\b(?:apartment|apt|suite|ste)\.?(?=\s|#)/i,
      word=>postalSpelling(postalUnits,word,modes.unit));
    if(field==='street2'||field==='billingStreet2'){
      const leading=unit?text.slice(0,unit.index):'';
      return unit&&!leading.trim()?leading+formattedUnit+trailing:value;
    }
    const street=unit?text.slice(0,unit.index):text;
    const parts=street.match(/^(\s*\d+[a-z]?(?:-\d+[a-z]?)?(?:\s+\d+\/\d+)?\s+)(.+?)(\s*,?\s*)$/i);
    if(!parts)return value;
    const words=parts[2].split(/(\s+)/),end=words.length-1;
    let suffix=end;
    if(end>=4&&postalLookup(postalDirections,words[end])&&postalLookup(postalSuffixes,words[end-2])){
      suffix=end-2;
      words[end]=postalSpelling(postalDirections,words[end],modes.direction);
    }
    if(suffix>=2&&postalLookup(postalSuffixes,words[suffix])){
      words[suffix]=postalSpelling(postalSuffixes,words[suffix],modes.suffix);
      if(suffix>=4&&!postalLookup(postalDirections,words[2]))words[0]=postalSpelling(postalDirections,words[0],modes.direction);
    }
    return parts[1]+words.join('')+parts[3]+formattedUnit+trailing;
  }
  function formatProfileCell(row,field,modes){
    const before=row[field]||'',group=profileFormatFields.get(field);
    if(group==='name'&&modes.name)return before.trim().replace(/\s+/g,' ');
    const country=field.startsWith('billing')?row.billingCountry:row.country;
    if(!isUSAddress(country))return before;
    if(group==='street')return formatPostalStreet(before,field,modes);
    if(group==='phone'&&modes.phone){
      const match=before.match(/^\s*(?:\+1[ .-]*)?(?:\((\d{3})\)|(\d{3}))[ .-]*(\d{3})[ .-]*(\d{4})\s*$/);
      if(match){
        const area=match[1]||match[2];
        return '('+area+') '+match[3]+'-'+match[4];
      }
    }
    return before;
  }
  function profileFormatPlan(){
    const modes=Object.fromEntries(Object.keys(profileFormatOptions).map(option=>[option,get('ec-format-'+option).checked]));
    const changes=new Map(),kept=new Set(),reserve=new Map();
    if(currentList()!==formatOwner)return changes;
    const shippingSelected=[...formatSelection.values()].some(fields=>shippingStreetFields.some(field=>fields.has(field)));
    for(const row of sheetRows){
      const key=sheetRowKey(row),edit={};
      for(const field of formatSelection.get(key)||[]){
        if(!profileFormatFields.has(field))continue;
        const after=formatProfileCell(row,field,modes);
        if((row[field]||'')!==after)edit[field]=after;
      }
      if(Object.keys(edit).length)changes.set(key,{row,edit});
      if(!shippingSelected)continue;
      const original=streetLines(row.shippingSource);
      if(!shippingStreetFields.every(field=>(row[field]||'')===original[field]))continue;
      if(!shippingStreetFields.some(field=>Object.hasOwn(edit,field)))kept.add(row.addressId);
      else if(!reserve.has(row.addressId))reserve.set(row.addressId,key);
    }
    for(const [destination,key] of reserve){
      if(kept.has(destination))continue;
      const {edit}=changes.get(key);
      for(const field of shippingStreetFields)delete edit[field];
      if(!Object.keys(edit).length)changes.delete(key);
    }
    return changes;
  }
  function positionProfileFormatDialog(){
    if(!formatDialog.open)return;
    const viewport=window.visualViewport,gap=12;
    const left=(viewport?.offsetLeft??0)+gap,top=(viewport?.offsetTop??0)+gap;
    const availableWidth=(viewport?.width??innerWidth)-2*gap,availableHeight=(viewport?.height??innerHeight)-2*gap;
    formatDialog.style.maxWidth=availableWidth+'px';
    formatDialog.style.maxHeight=availableHeight+'px';
    const width=formatDialog.offsetWidth,height=formatDialog.offsetHeight;
    const right=left+availableWidth,bottom=top+availableHeight;
    const anchor=formatButton.getBoundingClientRect();
    const clamp=(value,min,max)=>Math.max(min,Math.min(value,max));
    let x=clamp(anchor.left,left,right-width),y=clamp(anchor.top,top,bottom-height);
    if(anchor.right+gap+width<=right)x=anchor.right+gap;
    else if(anchor.left-gap-width>=left)x=anchor.left-gap-width;
    else if(anchor.bottom+gap+height<=bottom)y=anchor.bottom+gap;
    else if(anchor.top-gap-height>=top)y=anchor.top-gap-height;
    else y=top+(availableHeight-height)/2;
    formatDialog.style.left=x+'px';
    formatDialog.style.top=y+'px';
  }
  window.addEventListener('resize',positionProfileFormatDialog);
  window.visualViewport?.addEventListener('resize',positionProfileFormatDialog);
  function clearProfileFormatPreview(){
    for(const body of [...formatPreview.tBodies])body.remove();
    get('ec-format-results').hidden=true;
  }
  function showProfileFormatPreview(){
    const changes=profileFormatPlan(),fields=new Set();
    let fieldCount=0;
    for(const {edit} of changes.values()){
      for(const field of Object.keys(edit)){fields.add(field);fieldCount++;}
    }
    clearProfileFormatPreview();
    let shown=0;
    for(const {row,edit} of changes.values()){
      if(shown>=12)break;
      const body=formatPreview.createTBody();
      const title=document.createElement('th');
      title.colSpan=2;title.scope='rowgroup';title.textContent=row.nickname;
      body.insertRow().append(title);
      for(const [field,after] of Object.entries(edit)){
        if(shown>=12)break;
        if(fields.size>1){
          const caption=body.insertRow().insertCell();
          caption.colSpan=2;caption.className='ec-format-field';caption.textContent=fieldLabels.get(field);
        }
        const line=body.insertRow();
        line.insertCell().textContent=row[field]??'';
        line.insertCell().textContent=after;
        shown++;
      }
    }
    if(shown)get('ec-format-results').hidden=false;
    const summary=changes.size?changes.size+(changes.size===1?' profile will change.':' profiles will change.'):'No changes for the selected cells.';
    get('ec-format-status').textContent=summary+(fieldCount>shown?' Showing '+shown+' of '+fieldCount+' changed cells.':'');
    get('ec-format-apply').disabled=!changes.size;
    positionProfileFormatDialog();
  }
  formatButton.addEventListener('click',()=>{
    formatOwner=currentList();if(!formatOwner)return;
    formatSelection=new Map([...selectedCells].map(([key,fields])=>[key,new Set(fields)]));
    const fields=new Set([...formatSelection.values()].flatMap(selected=>[...selected]));
    for(const [option,appliesTo] of Object.entries(profileFormatOptions)){
      get('ec-format-'+option).closest('label').hidden=!appliesTo.some(field=>fields.has(field));
    }
    get('ec-format-preserve').hidden=!shippingStreetFields.some(field=>fields.has(field));
    showProfileFormatPreview();
    formatDialog.showModal();
    positionProfileFormatDialog();
  });
  for(const option of Object.keys(profileFormatOptions))get('ec-format-'+option).addEventListener('change',showProfileFormatPreview);
  get('ec-format-cancel').addEventListener('click',()=>formatDialog.close());
  formatDialog.addEventListener('close',()=>{formatOwner=null;formatSelection.clear();clearProfileFormatPreview();});
  get('ec-format-apply').addEventListener('click',()=>{
    const list=formatOwner,changes=profileFormatPlan();
    if(!changes.size){showProfileFormatPreview();return;}
    trackSheetEdit(list,[...changes.values()].map(({row})=>row.draftKey??row.accountId),()=>{
      for(const {row,edit} of changes.values())applySheetEdits(list,row,edit);
    });
    const selection=new Map(formatSelection),wasSelecting=selectingCells,point=formatPoint;
    formatDialog.close();
    redrawSheetEdits(list);
    for(const key of selection.keys())if(!sheetRowsByKey.has(key))selection.delete(key);
    selectingCells=wasSelecting;formatPoint=point;
    setSheetSelection(selection,[...selection.values()].reduce((count,fields)=>count+fields.size,0));
  });

  function syncRowInputs(tr,row,typing){
    let grew=false;
    for(const input of tr.querySelectorAll('input[data-field]')){
      const key=input.dataset.field,value=row[key]??'';
      if(input.type==='checkbox'){input.checked=value===true;continue;}
      if(input!==typing&&input.value===value)continue;
      if(input!==typing)input.value=value;
      sheetMeasure.font=measureFont(input);
      const width=Math.ceil(sheetMeasure.measureText(input.value).width+26);
      if(width>(fittedWidths.get(key)??0)){fittedWidths.set(key,width);grew=true;}
    }
    if(grew)applyColumnWidths();
  }
  function applyBulk(matchBilling){
    const list=currentList(),value=bulkValue.value;
    if(!matchBilling&&!/^(yes|no|true|false|1|0)$/i.test(value.trim())
        &&[...selectedCells.values()].some(fields=>[...flagFields].some(field=>fields.has(field)))){
      bulkValue.setCustomValidity('Enter yes or no for a checkbox field.');bulkValue.reportValidity();return;
    }
    bulkValue.setCustomValidity('');
    const byRow=[];
    for(const [key,fields] of selectedCells){
      const row=sheetRowsByKey.get(key);
      if(row)byRow.push([row,fields]);
    }
    trackSheetEdit(list,byRow.map(([row])=>row.draftKey??row.accountId),()=>{
      for(const [row,fields] of byRow){
        if(matchBilling)matchSheetBilling(list,row);
        else applySheetEdits(list,row,Object.fromEntries([...fields].map(field=>[field,value])));
      }
    });
    renderRows();saveCoordinatorChange({immediate:true});sheetScroller.focus({preventScroll:true});
  }
  get('ec-apply-cells').addEventListener('click',()=>applyBulk(false));
  bulkValue.addEventListener('input',()=>bulkValue.setCustomValidity(''));
  get('ec-match-billing').addEventListener('click',()=>applyBulk(true));
  function cancelBulk(){
    bulkValue.hidden=true;showSelection();sheetScroller.focus({preventScroll:true});
  }
  get('ec-cancel-cells').addEventListener('click',cancelBulk);
  bulkValue.addEventListener('keydown',event=>{
    if(event.isComposing)return;
    if(event.key==='Enter'){event.preventDefault();applyBulk(false);}
    else if(event.key==='Escape'){event.preventDefault();cancelBulk();}
  });
  sheetScroller.addEventListener('keydown',event=>{
    if(event.isComposing||event.target.closest('th'))return;
    if((event.ctrlKey||event.metaKey)&&!event.altKey){
      const key=event.key.toLowerCase();
      if(key==='z'||key==='y'){
        event.preventDefault();
        replaySheetEdit(key==='y'||event.shiftKey?'redo':'undo');
        return;
      }
    }
    const input=event.target.closest('input[data-row]');
    if(event.key==='Escape'){resetSheetSelection();return;}
    if(event.key==='Tab'&&input&&!event.ctrlKey&&!event.metaKey&&!event.altKey){
      const tr=input.closest('tr'),fields=visibleColumns.filter(([key])=>editableFields.has(key));
      const [first]=fields[0],[last]=fields.at(-1);
      if(input.dataset.field===(event.shiftKey?first:last)){
        const index=Number(tr.dataset.sheetIndex)+(event.shiftKey?-1:1);
        const column=visibleColumns.findIndex(([key])=>key===(event.shiftKey?last:first));
        const next=sheetInputAt(index,column);
        if(next){event.preventDefault();next.focus();next.select();}
      }
      return;
    }
    if(event.key==='Enter'){
      event.preventDefault();
      if(selectedCellCount>1||selectingCells){startBulk();return;}
      if(!input)return;
      const tr=input.closest('tr'),row=sheetRowsByKey.get(input.dataset.row);
      if(row)syncRowInputs(tr,row);
      saveCoordinatorChange({immediate:true});fitSheetColumns();
      const next=Number(tr.dataset.sheetIndex)+(event.shiftKey?-1:1);
      const nextInput=sheetInputAt(next,input.parentElement.cellIndex);
      if(nextInput){nextInput.focus();nextInput.select();}else{input.blur();}
      return;
    }
    if(selectedCellCount>1&&event.key.length===1&&!event.ctrlKey&&!event.metaKey&&!event.altKey){
      event.preventDefault();startBulk(event.key);
    }
  });
  sheetBody.addEventListener('input',event=>{
    const input=event.target,field=input.dataset.field,row=sheetRowsByKey.get(input.dataset.row);
    if(!row||!editableFields.has(field))return;
    const list=currentList();
    trackSheetEdit(list,[row.draftKey??row.accountId],()=>{
      applySheetEdits(list,row,{[field]:input.type==='checkbox'?input.checked:input.value});
    },input.type==='checkbox'?null:input);
    const tr=input.closest('tr');
    syncRowInputs(tr,row,input);
    tr.dataset.edited='true';
    if(field==='nickname')tr.dataset.moved='true';
  });
  sheetBody.addEventListener('focusout',event=>{
    const input=event.target,tr=input.closest('tr'),row=sheetRowsByKey.get(input.dataset.row);
    if(typingAction?.input===input)typingAction=null;
    if(row&&tr?.hasAttribute('data-edited')){
      syncRowInputs(tr,row);
      if(!tr.contains(event.relatedTarget)){delete tr.dataset.edited;fitSheetColumns();}
    }
    if(!tr?.hasAttribute('data-moved')||tr.contains(event.relatedTarget))return;
    delete tr.dataset.moved;
    resortSheetRows();
  });
  let selectedPerson='p-ros',selectedAddress='ros-home',sourceMode='address',sourceSerial=0,sourceTab='people';
  const addressFields=[['label','Address label (optional)'],['street','Street line 1'],['street2','Street line 2'],['city','City'],['state','State / region'],['zip','Postal code'],['country','Country'],['email','Contact email'],['phone','Contact phone']];
  const normalized=text=>String(text??'').normalize('NFKC').trim().toLowerCase();
  function isUSAddress(country){
    return ['us','usa','united states','united states of america'].includes(normalized(country).replace(/\./g,'').replace(/\s+/g,' '));
  }
  function splitAddressPostal(country,postal){
    const value=String(postal||'').trim(),match=isUSAddress(country)&&value.match(/^(\d{5})[- ]?(\d{4})$/);
    return match?{base:match[1],extension:match[2]}:{base:value,extension:''};
  }
  function joinAddressPostal(base,extension){return extension?base+'-'+extension:base;}
  function formPostal(form){
    return joinAddressPostal(form.elements.zip.value.trim(),form.elements.zip4.value.trim());
  }
  function postalInputs(record){
    const {base,extension}=splitAddressPostal(record.country,record.zip);
    return '<div class="ec-postal-fields"><label><span data-postal-label>Postal code</span><input name="zip" type="text" autocomplete="off" value="'+esc(base)+'"></label><label hidden>+4 (optional)<input type="text" name="zip4" inputmode="numeric" pattern="[0-9]{4}" title="Enter four digits." autocomplete="off" value="'+esc(extension)+'"></label></div>';
  }
  function updatePostalInputs(form){
    const postal=form.elements.zip,extra=form.elements.zip4,country=form.elements.country.value;
    const domestic=!country.trim()||isUSAddress(country);
    if(domestic){
      const {base,extension}=splitAddressPostal('US',postal.value);
      if(extension){postal.value=base;extra.value=extension;}
      postal.pattern='[0-9]{5}';postal.inputMode='numeric';postal.title='Enter a five-digit ZIP code.';
    }else{
      if(extra.value){postal.value=formPostal(form);extra.value='';}
      postal.removeAttribute('pattern');postal.removeAttribute('inputmode');postal.removeAttribute('title');
    }
    postal.required=domestic&&!!extra.value;
    extra.closest('label').hidden=!domestic;
    form.querySelector('[data-postal-label]').textContent=domestic?'ZIP code':'Postal code';
  }
  const ownerOf=id=>people.find(person=>person.id===id);
  const sourceText=(message='')=>{get('ec-source-status').textContent=message;};
  const directoryViews=['people','accounts','cards'];
  function requestedView(){
    const name=location.hash.slice(1);
    return directoryViews.includes(name)||name==='settings'&&!get('ec-show-settings').hidden?name:'coordinator';
  }
  function workView(view){
    const directory=directoryViews.includes(view),settings=view==='settings';
    if(sourceTab!==view||!directory)clearPrivateDirectory();
    get('ec-coordinator-view').hidden=directory||settings;get('ec-sources').hidden=!directory;
    get('ec-app-settings').hidden=!settings;
    get('ec-show-settings').setAttribute('aria-pressed',String(settings));
    get('ec-show-coordinator').setAttribute('aria-pressed',String(!directory&&!settings));
    root.dispatchEvent(new CustomEvent('ec:workspace-view',{detail:view}));
    for(const button of get('ec-source-tabs').querySelectorAll('[data-sources-tab]')){
      button.setAttribute('aria-pressed',String(button.dataset.sourcesTab===view));
    }
    if(settings)return;
    if(!directory){renderView(false);return;}
    sourceTab=view;
    for(const key of directoryViews)get('ec-settings-'+key).hidden=key!==view;
    if(view==='people')drawDirectory();
    root.dispatchEvent(new CustomEvent('ec:show-source',{detail:view}));
  }
  get('ec-source-tabs').addEventListener('click',event=>{
    const button=event.target.closest('button[data-sources-tab],#ec-show-coordinator,#ec-show-settings');if(!button)return;
    const view=button.dataset.sourcesTab||(button.id==='ec-show-settings'?'settings':'coordinator');
    if(document.querySelector('meta[name="parcel-workspace"]')){
      const url=new URL(location.href);url.hash=view==='coordinator'?'':view;
      history.replaceState(history.state,'',url);
    }
    workView(view);
  });
  function addressColor(address){return({home:'#e3a6fd',office:'#85dfff',veronica:'#89efc0'})[address.color]||'hsl('+address.hue+' 75% 76%)';}
  let paletteSheet=null,previousPalette=[];
  function updatePalette(){
    const sheet=document.getElementById('ec-source-colors').sheet;
    if(!sheet)return;
    const palette=addresses.filter(address=>Number.isFinite(address.hue));
    if(sheet===paletteSheet&&palette.length===previousPalette.length&&
      palette.every(({color,hue},i)=>previousPalette[i][0]===color&&previousPalette[i][1]===hue))return;
    while(sheet.cssRules.length)sheet.deleteRule(sheet.cssRules.length-1);
    // The earliest address wins when two aliases share a class.
    for(let i=palette.length-1;i>=0;i--){
      const address=palette[i],cls=CSS.escape('ec-'+address.color),tint='hsl('+address.hue+' 30% 13%)',ink=addressColor(address);
      sheet.insertRule('#event-coordinator .ec-address-row i.'+cls+'{--address-marker:'+ink+'}',sheet.cssRules.length);
      sheet.insertRule('#event-coordinator .ec-chip.'+cls+'{--chip-ink:'+ink+'}',sheet.cssRules.length);
      sheet.insertRule('#event-coordinator .'+cls+'{background:'+tint+';--ec-rail:'+ink+';--ec-row-fill:'+tint+'}',sheet.cssRules.length);
    }
    previousPalette=palette.map(({color,hue})=>[color,hue]);paletteSheet=sheet;
  }
  function directoryGroups(){return directoryAddresses().some(address=>!address.personId)?[...people,{id:'__unassigned__',name:'Unassigned addresses',unassigned:true}]:people;}
  function drawAddressPicker(){
    const byPerson=new Map();
    for(const parent of directoryAddresses()){
      const [main,...suites]=shippingDestinationOptions(parent);
      const options=[main,...suites.sort(compareShippingDestinations)].filter(address=>address.active!==false);
      if(!options.length)continue;
      const owner=parent.personId||'__unassigned__';
      if(!byPerson.has(owner))byPerson.set(owner,[]);
      const rows=byPerson.get(owner),grouped=!!(parent.parcelVariants?.length||parent.shippingVariants?.length);
      if(grouped)rows.push(['data-destination-group',parent.id,addressName(parent)+' · all destinations','ec-child']);
      for(const address of options){
        const label=grouped?[address.id===parent.id?'Main address':'',streetAndUnit(address)]:[address.label,streetAndUnit(address)];
        rows.push(['data-address',address.id,[...new Set(label.filter(Boolean))].join(' · '),grouped?'ec-grandchild':'ec-child',parent.id]);
      }
    }
    const rows=[];
    for(const person of directoryGroups()){
      const children=byPerson.get(person.id);
      if(children){rows.push(['data-parent',person.id,person.name,'']);for(const row of children)rows.push(row);}
    }
    addressPicker.set(rows);
  }
  function directoryAddressRow(address){
    const key=esc(address.id),label=esc(addressName(address)),enabled=address.active!==false;
    const add=address.parcelMissing?'':treeAddButton('destination',address.id,[address.person,addressName(address)].filter(Boolean).join(' · '));
    return '<div class="ec-directory-address"><button type="button" class="ec-address-row cursor-interaction" data-select-address="'+key+'" aria-current="'+(address.id===selectedAddress&&(sourceMode==='address'||sourceMode==='new-destination'))+'"><i aria-hidden="true" class="ec-'+esc(address.color)+'"></i>'+label+'</button><button type="button" role="switch" data-address-active="'+key+'" aria-label="'+esc([address.person,addressName(address),'active'].filter(Boolean).join(' '))+'" aria-checked="'+enabled+'">'+(enabled?'Active':'Inactive')+'</button>'+add+'</div>';
  }
  function treeAddButton(kind,id,parentName){
    const label=esc('Add '+kind+' under '+parentName);
    return '<button type="button" class="ec-directory-add cursor-interaction" data-add-'+kind+'-under="'+esc(id)+'" aria-label="'+label+'" title="'+label+'">+</button>';
  }
  function addressInputs(record,fields){
    return fields.map(([key,label])=>key==='zip'?postalInputs(record):'<label'+(key==='street'||key==='email'?' class="ec-source-long"':'')+'>'+label+'<input name="'+key+'" '+(['street','city','country'].includes(key)?'required ':'')+'autocomplete="off" value="'+esc(record[key]||'')+'"></label>').join('');
  }
  function destinationForm(parent){
    const record={...shippingLocation(parent),street2:''};
    const fields=addressFields.filter(([key])=>['street','street2','city','state','zip','country'].includes(key))
      .map(([key,label])=>[key,key==='street2'?'Suite, unit or street line 2':label]);
    const path=[parent.person,addressName(parent)].filter(Boolean).join(' · ');
    return '<span class="ec-source-kicker">'+esc(path)+'</span><h2>New shipping destination</h2><form id="ec-destination-form" class="ec-source-form" data-parent-address="'+esc(parent.id)+'"><div class="ec-source-fields">'+addressInputs(record,fields)+recipientEditor({})+'</div><p class="ec-source-note">Leave recipient names empty to use the names on '+esc(addressName(parent))+'.</p><label><input type="checkbox" name="enabled" checked> Use for shipping</label><div class="ec-source-actions"><button class="ec-primary cursor-interaction" type="submit">Add destination</button><button type="button" data-cancel-destination>Cancel</button></div></form>';
  }
  function recipientChip(kind,name){
    const {fullName,firstName,lastName,explicit}=parseRecipientName(name);
    const label=explicit?[firstName,lastName].join(' | '):fullName;
    return '<span class="ec-chip"><button type="button" class="ec-chip-label" data-edit-recipient-name="'+esc(name)+'" aria-label="Edit recipient '+esc(fullName)+'" title="Edit recipient name">'+esc(label)+'</button>'+deleteControl(kind,name,fullName)+'</span>';
  }
  function recipientEntry(kind,names,id,label){
    const suite=kind==='suite-recipient',inputAttr=suite?' data-suite-recipient-input':' name="newRecipient"';
    return '<div class="ec-recipient-entry">'+names.map(name=>recipientChip(kind,name)).join('')+
      '<input id="'+esc(id)+'" type="text"'+inputAttr+' aria-label="'+esc(label)+'" aria-describedby="'+esc(id)+'-hint" autocomplete="off" placeholder="Add a full shipping name">'+
      '<button type="button" data-split-recipient title="Insert the name separator at the cursor">Split here</button>'+
      '<button type="button" '+(suite?'data-suite-recipient-add':'data-add-recipient')+' aria-label="Add recipient name">+</button></div>'+
      '<p class="ec-recipient-hint" id="'+esc(id)+'-hint">Use | between first and last names, or place the cursor and tap Split here.</p>';
  }
  function recipientEditor(record){
    return '<div class="ec-recipient-field"><label for="ec-add-recipient">Recipient names</label>'+recipientEntry('recipient',addressRecipients(record),'ec-add-recipient','Recipient names')+'</div>';
  }
  function recipientInputError(input){
    const parsed=parseRecipientName(input.value);
    return parsed.explicit&&(!parsed.valid||!parsed.firstName||!parsed.lastName)?'Use one | with a first name on the left and a last name on the right.':'';
  }
  function updateRecipientHint(input){
    const hint=document.getElementById(input.id+'-hint'),error=recipientInputError(input);
    input.setCustomValidity(error);
    if(!hint)return;
    const {firstName,lastName}=parseRecipientName(input.value);
    hint.textContent=error||(input.value.trim()?'First name · '+(firstName||'—')+'    Last name · '+(lastName||'—'):'Use | between first and last names, or place the cursor and tap Split here.');
  }
  function clearRecipientInput(input){
    delete input.dataset.editingRecipient;input.value='';updateRecipientHint(input);
  }
  function commitRecipientInput(input,names,kind){
    const text=input.value.trim(),parsed=parseRecipientName(text);
    if(!text)return names;
    updateRecipientHint(input);
    if(!input.reportValidity())return null;
    const name=parsed.explicit?parsed.firstName+' | '+parsed.lastName:text;
    const index=names.indexOf(input.dataset.editingRecipient);
    let next;
    if(index>=0){
      if(names.some((other,i)=>i!==index&&recipientNameKey(other)===recipientNameKey(name))){
        input.setCustomValidity('This recipient is already listed.');input.reportValidity();return null;
      }
      next=names[index]===name?names:names.map((other,i)=>i===index?name:other);
    }else next=mergeRecipientName(names,name);
    if(next!==names){
      for(const chip of input.parentElement.querySelectorAll('.ec-chip'))chip.remove();
      input.insertAdjacentHTML('beforebegin',next.map(value=>recipientChip(kind,value)).join(''));
    }
    clearRecipientInput(input);input.focus();return next;
  }
  function recipientChips(form){
    return [...form.querySelectorAll('[data-remove-recipient]')].map(button=>button.dataset.removeRecipient);
  }
  function saveRecipientChips(form){
    const address=addresses.find(record=>record.id===form.dataset.sourceAddress);
    if(!address)return;
    setAddressDetails(address,{recipients:recipientChips(form)});
    refreshSourceRows(address.id);
    saveCoordinatorChange({immediate:true});
  }
  function addRecipient(form){
    const names=recipientChips(form);
    const next=commitRecipientInput(form.elements.namedItem('newRecipient'),names,'recipient');
    if(next&&next!==names)saveRecipientChips(form);
    return !!next;
  }
  function addressFormDetails(form){
    if(form.elements.namedItem('newRecipient').value.trim()&&!addRecipient(form))return null;
    return {recipients:recipientChips(form),email:form.elements.namedItem('email').value.trim(),phone:form.elements.namedItem('phone').value.trim()};
  }
  function updateAddressAction(form){
    const address=addresses.find(record=>record.id===form.dataset.sourceAddress);
    if(!address)return;
    const displayed={...address,...streetLines(address)};
    const dirty=['personId','label','street','street2','city','state','zip','country'].some(key=>{
      const field=form.elements.namedItem(key);
      if(key==='zip'){
        const {base,extension}=splitAddressPostal(address.country,address.zip);
        return !field.disabled&&formPostal(form)!==joinAddressPostal(base,extension);
      }
      return !field.disabled&&field.value.trim()!==(displayed[key]||'');
    });
    const action=form.querySelector('button[type=submit]');
    action.disabled=!dirty;action.hidden=!dirty;
  }
  function updateContact(input,immediate){
    const address=addresses.find(record=>record.id===input.form.dataset.sourceAddress);
    if(!address)return;
    const value=input.value.trim();
    if(value!==(address[input.name]||'')){
      setAddressDetails(address,{[input.name]:value});
      refreshSourceRows(address.id);
    }else if(!immediate)return;
    saveCoordinatorChange({immediate});
  }
  get('ec-source-editor').addEventListener('click',event=>{
    const edit=event.target.closest('[data-edit-recipient-name]'),split=event.target.closest('[data-split-recipient]');
    if(edit||split){
      const input=(edit||split).closest('.ec-recipient-entry').querySelector('input');
      if(edit){
        input.dataset.editingRecipient=edit.dataset.editRecipientName;
        input.value=edit.dataset.editRecipientName;input.focus();input.select();
      }else{
        const pos=input.selectionStart??input.value.length;
        const clean=part=>part.replaceAll('|',' ').trim();
        const first=clean(input.value.slice(0,pos)),last=clean(input.value.slice(pos));
        const left=first+(first?' ':'')+'| ';
        input.value=left+last;input.focus();input.setSelectionRange(left.length,left.length);
      }
      updateRecipientHint(input);
    }else if(event.target.closest('[data-cancel-destination]')){sourceMode='address';sourceText();drawDirectory(false);}
    else if(event.target.closest('[data-add-recipient]'))addRecipient(event.target.closest('form'));
  });
  get('ec-source-editor').addEventListener('keydown',event=>{
    if(event.key==='Escape'&&event.target.closest('.ec-recipient-entry')&&event.target.matches('input')){
      event.preventDefault();clearRecipientInput(event.target);return;
    }
    if(event.key!=='Enter'||event.isComposing)return;
    const input=event.target;
    if(input.name==='newRecipient'){event.preventDefault();addRecipient(input.form);}
    else if(input.form?.dataset.sourceAddress&&['email','phone'].includes(input.name)){event.preventDefault();updateContact(input,true);}
  });
  for(const type of ['input','change'])get('ec-source-editor').addEventListener(type,event=>{
    const input=event.target;
    if(input.matches('[name=newRecipient],[data-suite-recipient-input]')){
      if(!input.value)delete input.dataset.editingRecipient;
      updateRecipientHint(input);
    }
    if(input.form&&['ec-address-form','ec-destination-form'].includes(input.form.id)&&['zip','zip4','country'].includes(input.name))updatePostalInputs(input.form);
    if(!input.form?.dataset.sourceAddress)return;
    if(['email','phone'].includes(input.name))updateContact(input,type==='change');
    else if(input.name!=='newRecipient')updateAddressAction(input.form);
  });
  // Keep real, variable-height rows near the viewport. Measured placeholders preserve
  // scrolling when labels wrap or a suite's history and recipient editor expand.
  function mountPeopleChunks(host,records,{tag='div',size,estimate,paint,keep=()=>false,pinned}){
    const windowed=records.length>64;
    host.classList.toggle('ec-people-window',windowed);
    const chunks=[],parts=new Map(),top=host.scrollTop;
    const groupSize=windowed?size:Math.max(1,records.length);
    const fragment=document.createDocumentFragment();
    for(let start=0;start<records.length;start+=groupSize){
      const block=document.createElement(tag);
      const part=records.slice(start,start+groupSize);
      block.dataset.peopleChunk=String(chunks.length);
      if(windowed)block.style.height=(estimate*part.length)+'px';
      parts.set(block,part);chunks.push(block);fragment.append(block);
    }
    function mount(block){
      if(!block||block.firstElementChild)return;
      block.style.height='';
      paint(block,parts.get(block));
    }
    host.replaceChildren(fragment);
    mount(chunks[0]);
    if(!windowed)return ()=>{};
    if(pinned?.size)for(const block of chunks)if(parts.get(block).some(id=>pinned.has(id)))mount(block);
    host.scrollTop=top;
    const outside=new Set();
    function release(block){
      if(block?.firstElementChild&&!block.contains(document.activeElement)&&!keep(block)){
        block.style.height=block.getBoundingClientRect().height+'px';
        block.replaceChildren();
      }
    }
    const observer=new IntersectionObserver(entries=>{
      for(const {target,isIntersecting} of entries){
        if(isIntersecting){outside.delete(target);mount(target);}
        else{
          outside.add(target);release(target);
        }
      }
    },{root:host,rootMargin:'180px 0px'});
    for(const block of chunks)observer.observe(block);
    // Native Tab traversal can cross a chunk boundary before its intersection callback.
    function onTab(event){
      const block=event.target.closest('[data-people-chunk]');
      if(event.key==='Tab'&&block)mount(chunks[Number(block.dataset.peopleChunk)+(event.shiftKey?-1:1)]);
    }
    function onBlur(event){
      const block=event.target.closest('[data-people-chunk]');
      if(outside.has(block))queueMicrotask(()=>{if(outside.has(block))release(block);});
    }
    host.addEventListener('keydown',onTab);
    host.addEventListener('focusout',onBlur);
    return ()=>{observer.disconnect();outside.clear();host.removeEventListener('keydown',onTab);host.removeEventListener('focusout',onBlur);};
  }
  let stopPeopleTree=null,stopSuiteWindow=null;
  function drawPeopleTree(){
    stopPeopleTree?.();
    const groups=new Map(people.map(person=>[person.id,{person,addresses:[]}]));
    for(const address of directoryAddresses()){
      const owner=address.personId||'__unassigned__';
      if(!groups.has(owner)&&!address.personId)groups.set(owner,{person:{id:owner,name:'Unassigned addresses',unassigned:true},addresses:[]});
      groups.get(owner)?.addresses.push(address);
    }
    const rows=[];
    for(const group of groups.values()){
      rows.push({person:group.person,count:group.addresses.length});
      for(const address of group.addresses)rows.push({address});
    }
    stopPeopleTree=mountPeopleChunks(get('ec-source-list'),rows,{size:12,estimate:34,paint:(block,part)=>{
      block.innerHTML=part.map(({person,address,count})=>{
        if(address)return directoryAddressRow(address);
        const add=person.unassigned||person.parcelMissing?'':treeAddButton('address',person.id,person.name);
        return '<div class="ec-directory-person"><button type="button" class="ec-person-row cursor-interaction" data-select-person="'+esc(person.id)+'" '+(person.unassigned?'disabled':'')+' aria-current="'+(person.id===selectedPerson&&(sourceMode==='person'||sourceMode==='new-address'))+'">'+esc(person.name)+'<span>'+count+'</span></button>'+add+'</div>';
      }).join('');
    }});
  }
  function clearPeopleWindows(){
    stopPeopleTree?.();stopPeopleTree=null;
    stopSuiteWindow?.();stopSuiteWindow=null;
  }
  let renderedSourceMode=null,renderedSourceId=null;
  function drawDirectory(rebuildTree=true){
    if(get('ec-sources').hidden||get('ec-settings-people').hidden)return;
    const tree=get('ec-source-list');
    if(rebuildTree||!tree.firstElementChild)drawPeopleTree();
    else{
      tree.querySelector('[aria-current=true]')?.setAttribute('aria-current','false');
      const isAddress=sourceMode==='address'||sourceMode==='new-destination';
      const id=isAddress?selectedAddress:sourceMode==='new-person'?null:selectedPerson;
      if(id)tree.querySelector('[data-select-'+(isAddress?'address':'person')+'="'+CSS.escape(id)+'"]')?.setAttribute('aria-current','true');
    }
    stopSuiteWindow?.();stopSuiteWindow=null;
    const person=ownerOf(selectedPerson),address=addresses.find(item=>item.id===selectedAddress);
    const editor=get('ec-source-editor'),isPerson=sourceMode==='person'||sourceMode==='new-person',existing=sourceMode==='address'||sourceMode==='person';
    const recordId=sourceMode==='address'||sourceMode==='new-destination'?selectedAddress:selectedPerson;
    const sameRecord=sourceMode===renderedSourceMode&&recordId===renderedSourceId;
    const top=sameRecord?editor.scrollTop:0,suiteTop=sameRecord?editor.querySelector('.ec-destination-chunks')?.scrollTop:0;
    renderedSourceMode=sourceMode;renderedSourceId=recordId;
    function finishRender(){
      root.dispatchEvent(new Event('ec:directory-rendered'));
      const suites=editor.querySelector('.ec-destination-chunks');
      if(suites&&suiteTop)suites.scrollTop=suiteTop;
      editor.scrollTop=top;
      // Restore scrolling after the destination editor finishes laying out.
      if(editor.scrollTop<top){
        const clamped=editor.scrollTop,form=editor.querySelector('form');
        requestAnimationFrame(()=>{if(editor.contains(form)&&editor.scrollTop===clamped)editor.scrollTop=top;});
      }
    }
    if(sourceMode==='new-destination'&&address){
      editor.innerHTML=destinationForm(address);
      updatePostalInputs(get('ec-destination-form'));
      finishRender();
      return;
    }
    if(isPerson){
      const name=existing&&person?person.name:'';
      editor.innerHTML='<span class="ec-source-kicker">Person</span><h2>'+esc(name||'New person')+'</h2><form id="ec-person-form" class="ec-source-form"><div class="ec-source-fields"><label>Person nickname<input required name="name" autocomplete="off" value="'+esc(name)+'"></label></div><div class="ec-source-actions"><button class="ec-primary cursor-interaction" type="submit">'+(existing?'Update person':'Create person')+'</button></div></form>';
    }else{
      const record=existing&&address?address:{personId:selectedPerson};
      const owners=people;
      const options='<option value="">Unassigned</option>'+owners.map(item=>'<option value="'+esc(item.id)+'"'+(record.personId===item.id?' selected':'')+'>'+esc(item.name)+'</option>').join('');
      const fields=addressInputs({...record,...streetLines(record)},addressFields);
      const linked=existing?events.reduce((count,event)=>count+event.lists.reduce((sum,list)=>sum+list.rows.filter(row=>row.parentAddressId===record.id).length,0),0):0;
      const note=existing?'Recipient names and contact details update as you edit.':'Create this address to keep its recipients and contact details.';
      editor.innerHTML='<span class="ec-source-kicker">Address</span><h2>'+esc(existing?[record.person,addressName(record)].filter(Boolean).join(' · '):'New address')+'</h2><form id="ec-address-form" class="ec-source-form"'+(existing?' data-source-address="'+esc(record.id)+'"':'')+'><div class="ec-source-fields"><label>Person nickname<select name="personId" class="cursor-interaction">'+options+'</select></label>'+recipientEditor(record)+fields+'</div><p class="ec-source-note" data-address-note>'+note+'</p><div class="ec-source-actions"><button type="submit" class="ec-primary cursor-interaction"'+(existing?' hidden disabled':'')+'>'+(existing?'Update address':'Create address')+'</button></div></form>'+(existing?'<div class="ec-source-meta"><span><strong>'+linked+'</strong> linked '+(linked===1?'profile':'profiles')+' across events</span></div>':'');
    }
    if(sourceMode==='new-address'&&person)editor.querySelector('.ec-source-kicker').textContent=person.name+' · Address';
    if(!isPerson)updatePostalInputs(get('ec-address-form'));
    finishRender();
  }
  function freshAddress(person,values){
    const sequence=++sourceSerial;
    return {active:true,recipients:[],...values,id:'src-a-'+sequence,personId:person?.id||null,person:person?.name||'',color:'source-'+sequence,hue:(sequence*137+24)%360};
  }
  function afterMasterChange(message){
    drawAddressPicker();updatePalette();refreshSourceRows();renderConfig(false);drawDirectory();sourceText(message);
  }
  get('ec-source-editor').addEventListener('submit',event=>{
    event.preventDefault();const form=event.target,read=name=>String(form.elements.namedItem(name)?.value||'').trim();
    if(form.id==='ec-destination-form'){
      const parent=directoryAddresses().find(record=>record.id===form.dataset.parentAddress&&!record.parcelMissing);
      if(!parent||sourceMode!=='new-destination'||selectedAddress!==parent.id){sourceText('Select the address group again.');return;}
      const location=Object.fromEntries(importLocationFields.map(field=>[field,read(field)]));
      location.zip=formPostal(form);
      if(isUSAddress(location.country)&&/^\d[\dA-Za-z/-]*$/.test(location.street2))location.street2='Suite '+location.street2;
      const fields=shippingLocation(location);
      if(!fields.street||!fields.city||!fields.country){sourceText('Enter a street, city and country.');return;}
      const key=destinationLocationKey(fields);
      if([parent,...parent.parcelVariants||[],...parent.shippingVariants||[]].some(record=>destinationLocationKey(record)===key)){
        sourceText('This destination already exists under '+addressName(parent)+'.');return;
      }
      if(read('newRecipient')&&!addRecipient(form))return;
      const id=parent.id+':destination:'+(++sourceSerial);
      (parent.shippingVariants??=[]).push({id,authored:true,...fields,enabled:form.elements.enabled.checked,recipients:recipientChips(form)});
      if(form.elements.enabled.checked)reconcileShippingParent(parent);
      sourceMode='address';
      afterMasterChange('Added '+(fields.street2||fields.street)+' under '+addressName(parent)+'.');
      saveCoordinatorChange({immediate:true});return;
    }
    if(form.id==='ec-person-form'){
      const name=read('name'),adding=sourceMode==='new-person';if(!name){sourceText('Enter a display name.');return;}
      if(people.some(person=>normalized(person.name)===normalized(name)&&(adding||person.id!==selectedPerson))){sourceText('A person already uses that name.');return;}
      if(adding){const person={id:'src-p-'+(++sourceSerial),name};people.push(person);selectedPerson=person.id;}
      else{ownerOf(selectedPerson).name=name;for(const address of addresses)if(address.personId===selectedPerson)address.person=name;}
      sourceMode='person';selectedAddress=null;afterMasterChange(adding?'Person added. Choose + Address to add their first address.':'Person updated across events.');return;
    }
    if(form.id!=='ec-address-form')return;
    const person=ownerOf(read('personId')),adding=sourceMode==='new-address',values=Object.fromEntries(addressFields.map(([key])=>[key,read(key)]));
    values.zip=formPostal(form);
    const details=addressFormDetails(form);
    if(!details)return;
    Object.assign(values,details);
    if(['street','city','country'].some(key=>!values[key])){sourceText('Enter a street, city and country.');return;}
    values.label=addressName(values);
    if(addresses.some(address=>address.personId===(person?.id||null)&&normalized(address.label)===normalized(values.label)&&(adding||address.id!==selectedAddress))){sourceText('This person already has an address with that label.');return;}
    if(adding){const record=freshAddress(person,values);addresses.push(record);selectedAddress=record.id;}
    else{
      const record=addresses.find(item=>item.id===selectedAddress),displayed=streetLines(record);
      if(values.street===displayed.street&&values.street2===displayed.street2){
        values.street=record.street;values.street2=record.street2;
      }
      Object.assign(record,values,{personId:person?.id||null,person:person?.name||''});
    }
    selectedPerson=person?.id||null;sourceMode='address';afterMasterChange(adding?'Address added to the master directory.':'Linked profiles updated. Individual cell edits were kept.');
  });
  get('ec-source-list').addEventListener('click',event=>{
    const addAddress=event.target.closest('[data-add-address-under]'),addDestination=event.target.closest('[data-add-destination-under]');
    if(addAddress){startNewAddress(addAddress.dataset.addAddressUnder);return;}
    if(addDestination){startNewDestination(addDestination.dataset.addDestinationUnder);return;}
    const toggle=event.target.closest('[data-address-active]');
    if(toggle){
      const address=addresses.find(record=>record.id===toggle.dataset.addressActive);
      address.active=address.active===false;
      toggle.setAttribute('aria-checked',String(address.active));toggle.textContent=address.active?'Active':'Inactive';
      drawAddressPicker();renderConfig(false);saveCoordinatorChange();
      sourceText(addressName(address)+(address.active?' is available for new profiles.':' is inactive for new profiles.'));
      return;
    }
    const addressButton=event.target.closest('[data-select-address]'),personButton=event.target.closest('[data-select-person]');
    if((addressButton&&sourceMode==='address'&&addressButton.dataset.selectAddress===selectedAddress)||
      (personButton&&sourceMode==='person'&&personButton.dataset.selectPerson===selectedPerson))return;
    if(addressButton){selectedAddress=addressButton.dataset.selectAddress;selectedPerson=addresses.find(record=>record.id===selectedAddress).personId;sourceMode='address';}
    else if(personButton){selectedPerson=personButton.dataset.selectPerson;selectedAddress=null;sourceMode='person';}else return;
    sourceText();drawDirectory(false);
  });
  get('ec-new-person').addEventListener('click',()=>{sourceText();sourceMode='new-person';selectedAddress=null;drawDirectory(false);get('ec-person-form').elements.namedItem('name').focus();});
  function startNewAddress(personId){
    sourceText();sourceMode='new-address';selectedAddress=null;selectedPerson=personId;drawDirectory(false);
    get('ec-address-form').elements.namedItem('label').focus();
  }
  function startNewDestination(parentId){
    const parent=directoryAddresses().find(record=>record.id===parentId&&!record.parcelMissing);
    if(!parent)return;
    selectedAddress=parent.id;selectedPerson=parent.personId;sourceMode='new-destination';sourceText();drawDirectory(false);
    get('ec-destination-form').elements.street2.focus();
  }
  get('ec-new-address').addEventListener('click',()=>startNewAddress(selectedPerson));
  const importLocationFields=['street','street2','city','state','zip','country'];
  const addressImportFields=[
    ['nickname','Person nickname',['person','person label','display name','owner','alias']],
    ['recipient','Recipient full name',['name','full name','recipient name','shipping name','ship to name','recipient names']],
    ['firstName','Recipient first name',['first name','shipping first name','given name']],
    ['lastName','Recipient last name',['last name','shipping last name','surname','family name']],
    ['label','Address label',['address nickname','location label']],
    ['street','Street line 1',['address','address1','address line 1','line1','shipping address','shipping street']],
    ['street2','Street line 2',['address2','address line 2','line2','apt','unit','shipping address 2','shipping street 2']],
    ['city','City',['town','shipping city']],
    ['state','State / region',['region','province','shipping state']],
    ['zip','Postal code',['zip code','zip5','postal','post code','postcode','shipping zip code','shipping post code']],
    ['zip4','ZIP extension (+4)',['+4','zip4','zip code +4','shipping zip4','zip plus four','zip plus 4','zip extension','postal code extension','zip suffix','plus 4']],
    ['country','Country',['country code','shipping country']],
    ['phone','Contact phone',['phone number','shipping phone','phone num']],
    ['email','Contact email',['email address','shipping email']],
    ['active','Active',['enabled','status','is active']]
  ].map(([key,label,aliases])=>({key,label,aliases}));
  function importLocationKey(person,location){
    const {base,extension}=splitAddressPostal(location.country,location.zip);
    return JSON.stringify([normalized(person),...importLocationFields.map(key=>normalized(key==='zip'?joinAddressPostal(base,extension):location[key]).replace(/\s+/g,' '))]);
  }
  function sameImportedSplit(left,right){
    return recipientNameKey(left.firstName)===recipientNameKey(right.firstName)&&recipientNameKey(left.lastName)===recipientNameKey(right.lastName);
  }
  function inspectAddressImport(rows){
    const issues=[],destinations=new Map(),formatted=new Map(),labels=new Map(),owners=new Map(),planned=new Set();
    const ownerKey=(person,label)=>JSON.stringify([normalized(person),normalized(label)]);
    function indexLocation(entry,location,snapshot=false){
      const coordinates=[[formatted,comparableLocation(location)]];
      if(!snapshot)coordinates.push([destinations,location]);
      for(const [index,value] of coordinates){
        const key=importLocationKey(entry.person,value);
        index.set(key,index.has(key)&&index.get(key)!==entry?null:entry);
      }
    }
    for(const person of people.filter(item=>!item.parcelMissing)){
      const key=normalized(person.name),matches=owners.get(key)||[];
      matches.push(person);owners.set(key,matches);
    }
    for(const record of directoryAddresses().filter(item=>!item.parcelMissing)){
      const person=ownerOf(record.personId)?.name;
      if(!person)continue;
      const entry={person,label:record.label,existing:record,details:{},recipients:new Map()};
      labels.set(ownerKey(person,record.label),entry);
      indexLocation(entry,record);
      const variants=new Map(),observed=record.parcelVariants||[];
      const variantId=shippingVariantLookup(record),observedIds=new Set(observed.map(variantId));
      for(const location of [...record.shippingVariants||[],...observed]){
        const id=location.id||variantId(location);
        if(!id)continue;
        let variant=variants.get(id);
        if(!variant){
          variant={person,label:record.label,existing:record,variantId:id,variantLocation:location,details:{},recipients:new Map()};
          variants.set(id,variant);
        }
        indexLocation(variant,location,observedIds.has(location.id));
      }
    }
    for(const row of rows){
      const read=key=>(row[key]||'').trim(),issue=message=>issues.push('Row '+row.__row+' · '+message);
      const full=read('recipient'),firstName=read('firstName'),lastName=read('lastName'),hasParts=!!(firstName||lastName);
      const fullParts=parseRecipientName(full);
      if(full&&(!fullParts.valid||(fullParts.explicit&&(!fullParts.firstName||!fullParts.lastName)))){
        issue('Use one | between the first and last names, or map them as separate columns.');continue;
      }
      if(firstName.includes('|')||lastName.includes('|')){issue('Remove the | separator from the separately mapped name columns.');continue;}
      const parts=formatRecipientParts(firstName,lastName);
      if(full&&hasParts&&(recipientNameKey(full)!==recipientNameKey(parts)
          ||(fullParts.explicit&&!sameImportedSplit(fullParts,{firstName,lastName})))){
        issue('The full name and split recipient name disagree. Choose one mapping.');continue;
      }
      const recipient=fullParts.explicit?[fullParts.firstName,fullParts.lastName].join(' | '):hasParts?parts:full;
      const parsed=parseRecipientName(recipient),explicit=fullParts.explicit||hasParts;
      const person=read('nickname')||parsed.fullName,explicitLabel=read('label');
      const location=Object.fromEntries(importLocationFields.map(key=>[key,read(key)]));
      if(!person){issue('Provide a person nickname or recipient name.');continue;}
      const missing=['street','city','country'].filter(key=>!location[key]);
      if(missing.length){issue('Missing '+missing.join(', ')+'.');continue;}
      let extra=read('zip4');
      if(isUSAddress(location.country)){
        if(!location.zip&&splitAddressPostal(location.country,extra).extension){location.zip=extra;extra='';}
        const {base,extension}=splitAddressPostal(location.country,location.zip);
        if((base||extra)&&!/^\d{5}$/.test(base)){issue('Enter a five-digit ZIP code, optionally followed by four digits.');continue;}
        if(extra&&!/^\d{4}$/.test(extra)){issue('The ZIP extension must be four digits.');continue;}
        if(extra&&extension&&extra!==extension){issue('The postal code and ZIP extension disagree.');continue;}
        location.zip=joinAddressPostal(base,extra||extension);
      }else if(extra){issue('Use the Postal code column for international addresses and leave the US ZIP extension empty.');continue;}
      if((owners.get(normalized(person))||[]).length>1){issue('More than one person has this nickname. Select the person in the directory.');continue;}
      const details={};
      for(const key of ['email','phone'])if(read(key))details[key]=read(key);
      if(read('active')){
        const value=normalized(row.active);
        if(['true','yes','1','active','enabled','on'].includes(value))details.active=true;
        else if(['false','no','0','inactive','disabled','off'].includes(value))details.active=false;
        else {issue('Use active/inactive, yes/no, true/false or 1/0 for Active.');continue;}
      }
      const key=importLocationKey(person,location);
      let entry=destinations.has(key)?destinations.get(key):formatted.get(importLocationKey(person,comparableLocation(location)));
      if(entry===null){issue('This destination matches multiple saved addresses. Resolve the match in the directory.');continue;}
      if(!entry){
        const stem=streetAndUnit(location).slice(0,200);
        let label=explicitLabel||stem;
        if(explicitLabel&&labels.has(ownerKey(person,label))){issue('This address label belongs to a different destination for the same person.');continue;}
        for(let suffix=2;labels.has(ownerKey(person,label));suffix++){
          const tail=' ('+suffix+')';label=stem.slice(0,200-tail.length)+tail;
        }
        entry={person,label,location,owner:(owners.get(normalized(person))||[])[0],explicitLabel,details:{},recipients:new Map()};
        indexLocation(entry,location);labels.set(ownerKey(person,label),entry);
      }else if(explicitLabel&&normalized(entry.label)!==normalized(explicitLabel)){
        if(entry.existing||entry.explicitLabel||labels.has(ownerKey(person,explicitLabel))){issue('The same destination already has a different label. Leave the label blank or use its existing label.');continue;}
        labels.delete(ownerKey(person,entry.label));entry.label=explicitLabel;labels.set(ownerKey(person,explicitLabel),entry);
      }
      if(explicitLabel)entry.explicitLabel=explicitLabel;
      if(Object.keys(details).some(key=>key in entry.details&&entry.details[key]!==details[key])){issue('Rows for this address disagree on contact details or active status.');continue;}
      Object.assign(entry.details,details);
      if(recipient){
        const identity=recipientNameKey(recipient),prior=entry.recipients.get(identity);
        if(prior?.explicit&&explicit&&!sameImportedSplit(parseRecipientName(prior.name),parsed)){
          issue('Rows for this destination split the same recipient name differently. Choose one split.');continue;
        }
        if(!prior||explicit&&!prior.explicit)entry.recipients.set(identity,{name:recipient,explicit});
      }
      planned.add(entry);
    }
    const changes=[];let added=0,updated=0,unchanged=0;
    for(const entry of planned){
      const prior=entry.variantId?(entry.existing.shippingVariants||[]).find(variant=>variant.id===entry.variantId):entry.existing;
      const details={...entry.details},names=prior?addressRecipients(prior):[];
      let merged=names;
      for(const {name,explicit} of entry.recipients.values())merged=mergeRecipientName(merged,name,explicit);
      if(!entry.existing||names.length!==merged.length||names.some((name,index)=>name!==merged[index]))details.recipients=merged;
      if(entry.existing){
        const active=entry.variantId?!!prior?.enabled:prior.active!==false;
        for(const key of ['email','phone','active'])if(details[key]===(key==='active'?active:prior?.[key]||''))delete details[key];
        if(!Object.keys(details).length){unchanged++;continue;}
        updated++;
      }else added++;
      changes.push({...entry,details});
    }
    const summary=added+' new · '+updated+' updated · '+unchanged+' unchanged';
    return {issues,summary,actions:issues.length?0:changes.length,apply:()=>applyImportedAddresses(changes,summary)};
  }
  function selectImportedAddress(address){
    if(address){selectedAddress=address.id;selectedPerson=address.personId;sourceMode='address';}
  }
  function setImportedAddressDetails(address,change){
    if(!change.variantId){setAddressDetails(address,change.details);return;}
    const variant=(address.shippingVariants||[]).find(item=>item.id===change.variantId)||ensureShippingVariant(address,change.variantLocation);
    const {active,...metadata}=change.details;
    Object.assign(variant,metadata);
    if(typeof active==='boolean'&&variant.enabled!==active){variant.enabled=active;reconcileShippingParent(address);}
  }
  let applyImportedAddresses=async(changes,summary)=>{
    const owners=new Map(people.filter(person=>!person.parcelMissing).map(person=>[normalized(person.name),person]));
    let first=null;
    for(const change of changes){
      let address=change.existing;
      if(!address){
        const key=normalized(change.person);
        let person=owners.get(key);
        if(!person){person={id:'src-p-'+(++sourceSerial),name:change.person};people.push(person);owners.set(key,person);}
        address=freshAddress(person,{...change.location,label:change.label});
        addresses.push(address);
      }
      setImportedAddressDetails(address,change);
      if(!first)first=address;
    }
    selectImportedAddress(first);
    afterMasterChange('Address import · '+summary+'.');
  };
  mappedImportAdapters.addresses={fields:addressImportFields,title:()=> 'Import people & addresses',inspect:inspectAddressImport};
  get('ec-import-toggle').addEventListener('click',()=>openMappedImport('addresses'));
  drawAddressPicker();
  function restore(snapshot){
    const view=snapshot?.modelContent;if(view?.preview!=='event-coordinator')return;
    const event=events.find(event=>event.id===view.eventId)||events[view.event];
    if(event){activeEvent=event.id;const list=event.lists.find(list=>list.id===(view.listId||view.bot));if(list)event.selectedListId=list.id;}
  }
  const inv=name=>get('ec-inv-'+name);
  let editedAccount=0,editedCard='',credentialExportText='';
  let accountScope={retailer:null,group:null},accountSort='email',accountDirection=1,credentialView=0;
  const checkedAccounts=new Set(),expandedAccountEmails=new Set(),emailAccountGroups=new Map();
  let readRetailPasswords=async records=>Object.fromEntries(records.map(account=>[account.id,account.password||'']));
  let readDirectorySecrets=async(kind,records)=>kind==='cards'?resolveCheckout(records.map(record=>record.id)):readRetailPasswords(records);
  let reloadDirectorySources=()=>{},directoryHasPendingChanges=()=>false;
  let forgetAccountSecrets=()=>{},accountEditorChanged=()=>{};
  const directorySecrets=new Map(),pendingDirectoryIds=new Set(),failedDirectoryIds=new Set();
  let directoryGeneration=0,directoryTimer=0;
  function invalidateDirectorySecrets(){
    directoryGeneration++;clearTimeout(directoryTimer);directorySecrets.clear();pendingDirectoryIds.clear();failedDirectoryIds.clear();
  }
  function clearPrivateDirectory(){
    invalidateDirectorySecrets();
    directoryLogins=[];directoryItems=[];directoryAccountsById.clear();directoryCards=[];emailAccountGroups.clear();
    // Each directory rebuilds its list from the current records when opened.
    clearPeopleWindows();get('ec-source-list').replaceChildren();get('ec-source-editor').replaceChildren();
    inv('account-list').replaceChildren();inv('card-list').replaceChildren();
    for(const field of [inv('card-number'),inv('card-code'),inv('password')]){field.value='';delete field.dataset.edited;}
    resetTransfer();
  }
  function fillPrivateField(field,value){
    if(field.dataset.edited)return;
    field.value=value||'';
    if(field===inv('card-number'))checkEditedCardNumber();
  }
  function fillDirectoryEditor(kind){
    if(sourceTab!==kind||get('ec-sources').hidden)return;
    if(kind==='accounts'&&!inv('account-form').hidden){
      const detail=directorySecrets.get(String(editedAccount));
      if(detail)fillPrivateField(inv('password'),detail.password);
    }else if(kind==='cards'&&!inv('card-form').hidden){
      const detail=directorySecrets.get(cardIds.get(editedCard));
      if(detail){fillPrivateField(inv('card-number'),detail.number);fillPrivateField(inv('card-code'),detail.cvv);}
    }
  }
  function paintDirectory(kind){
    const body=inv(kind==='cards'?'card-list':'account-list');
    for(const cell of body.querySelectorAll('[data-secret-cell]')){
      const detail=directorySecrets.get(cell.dataset.secretId);
      if(ecPrivacyEnabled())cell.textContent=cell.dataset.mask;
      else if(detail)cell.textContent=detail[cell.dataset.secretCell]||cell.dataset.fallback;
      else cell.textContent=failedDirectoryIds.has(cell.dataset.secretId)?'Unavailable':'Loading';
    }
    fillDirectoryEditor(kind);
  }
  function directoryReadError(kind,error,message){
    const action=document.createElement('button');
    action.type='button';action.dataset.retryDetails='';action.textContent='Retry loading details';
    if(error.status===409){
      action.textContent='Sync now';action.addEventListener('click',()=>reloadDirectorySources());
      message.replaceChildren(directoryHasPendingChanges()?'Another device saved changes. Finish your edit to sync. ':'Loading changes from another device. ',action);
      return;
    }
    action.addEventListener('click',()=>refreshDirectorySecrets(kind));
    if(error.status===401||error.status===403){
      message.replaceChildren((error.message||'Please reload this workspace.')+' ',action);
    }else message.replaceChildren((error.message||'Could not load saved details.')+' ',action);
  }
  function displayedDirectoryIds(kind){
    const cards=kind==='cards',body=inv(cards?'card-list':'account-list');
    const ids=new Set([...body.querySelectorAll(cards?'[data-card-record]':'[data-account-record]')]
      .map(row=>cards?row.dataset.cardRecord:row.dataset.accountRecord));
    const editorId=cards?cardIds.get(editedCard):editedAccount;
    if(editorId&&!inv(cards?'card-form':'account-form').hidden)ids.add(String(editorId));
    return ids;
  }
  function refreshDirectorySecrets(kind){
    if(sourceTab!==kind||get('ec-sources').hidden)return;
    paintDirectory(kind);
    if(ecPrivacyEnabled())return;
    const message=inv(kind==='cards'?'card-msg':'account-msg');
    if(message.querySelector('[data-retry-details]'))message.replaceChildren();
    const visible=displayedDirectoryIds(kind);
    const records=Array.from(visible,id=>kind==='cards'?{id}:directoryAccountsById.get(Number(id))).filter(Boolean);
    for(const id of directorySecrets.keys())if(!visible.has(id))directorySecrets.delete(id);
    clearTimeout(directoryTimer);
    const generation=directoryGeneration;
    directoryTimer=setTimeout(async()=>{
      const missing=records.filter(r=>!directorySecrets.has(String(r.id))&&!pendingDirectoryIds.has(String(r.id)));
      for(const record of missing){pendingDirectoryIds.add(String(record.id));failedDirectoryIds.delete(String(record.id));}
      paintDirectory(kind);
      try{
        for(let start=0;start<missing.length;start+=1000){
          const batch=missing.slice(start,start+1000);
          if(generation!==directoryGeneration)return;
          const values=await readDirectorySecrets(kind,batch);
          if(generation!==directoryGeneration||ecPrivacyEnabled())return;
          const shown=displayedDirectoryIds(kind);
          for(const {id} of batch){
            const key=String(id);if(!shown.has(key))continue;
            if(kind==='accounts')directorySecrets.set(key,{password:values[id]||''});
            else{const card=values.get(id)||{};directorySecrets.set(key,{number:card.number||'',cvv:card.cvv||''});}
          }
          paintDirectory(kind);
        }
      }catch(error){
        if(generation!==directoryGeneration||ecPrivacyEnabled())return;
        for(const record of missing)if(!directorySecrets.has(String(record.id)))failedDirectoryIds.add(String(record.id));
        paintDirectory(kind);
        directoryReadError(kind,error,message);
      }
      finally{if(generation===directoryGeneration)for(const record of missing)pendingDirectoryIds.delete(String(record.id));}
    },40);
  }
  function secretCell(id,field,mask,fallback='Not set'){
    return '<td data-secret-id="'+esc(String(id))+'" data-secret-cell="'+field+'" data-mask="'+esc(mask)+'" data-fallback="'+esc(fallback)+'">'+esc(ecPrivacyEnabled()?mask:fallback)+'</td>';
  }
  for(const field of [inv('card-number'),inv('card-code'),inv('password')])field.addEventListener('input',()=>{
    field.dataset.edited='true';field.setCustomValidity('');
  });
  function showCardNumberIssue(message,block=Boolean(message)){
    const input=inv('card-number');
    input.setCustomValidity(block?message:'');
    if(block)input.setAttribute('aria-invalid','true');else input.removeAttribute('aria-invalid');
    inv('pan-error').textContent=message;
  }
  function checkEditedCardNumber(){
    const input=inv('card-number'),digits=input.value.replace(/[\s-]/g,'');
    const message=input.value?cardNumberIssue(digits):'';
    const block=Boolean(message&&input.dataset.edited);
    showCardNumberIssue(message,block);
    return !block;
  }
  inv('card-number').addEventListener('input',()=>showCardNumberIssue(''));
  inv('card-number').addEventListener('blur',checkEditedCardNumber);
  root.addEventListener('ec:card-editor',()=>showCardNumberIssue(''));
  inv('card-form').addEventListener('submit',event=>{
    if(checkEditedCardNumber())return;
    event.preventDefault();event.stopImmediatePropagation();inv('card-number').reportValidity();
  },true);
  const selectedRetailer=()=>inv('retailer').value;
  const retailerShowsUsername=retailer=>!['walmart','target'].includes(normalized(retailer));
  let accountUsernamesVisible=true;
  const selectedAccounts=()=>masterAccounts.filter(account=>account.retailer===selectedRetailer());
  const groupTitle=id=>accountGroups.find(group=>group.id===id)?.name||'';
  const retailerTitle=retailer=>retailer||'Unassigned';
  const retailerNames=()=>[...new Set([...supportedRetailers,...masterAccounts.filter(managedRetailAccount).map(a=>a.retailer),...accountGroups.map(g=>g.retailer)].filter(Boolean))].sort(nameCollator.compare);
  function retailOptions(){
    const options=retailerNames().map(name=>'<option>'+esc(name)+'</option>').join('');
    for(const select of [inv('retailer'),inv('bulk-retailer'),get('ec-list-retailer')]){
      const current=select.value||'Target';
      if(select.dataset.options!==options){select.innerHTML=options;select.dataset.options=options;}
      select.value=[...select.options].some(item=>item.value===current)?current:'Target';
    }
  }
  function visibleAccounts(){
    const query=normalized(inv('search').value);
    return masterAccounts.filter(account=>managedRetailAccount(account)&&(accountScope.retailer===null||account.retailer===accountScope.retailer)&&
      (!accountScope.group||account.groupIds?.includes(accountScope.group))&&
      (!query||[account.email,account.username,account.retailer].some(value=>normalized(value).includes(query))))
      .sort((a,b)=>accountDirection*nameCollator.compare(a[accountSort]||'',b[accountSort]||'')||a.id-b.id);
  }
  const scopeLabel=()=>accountScope.group?accountScope.retailer+' / '+groupTitle(accountScope.group):
    accountScope.retailer===null?'All accounts':retailerTitle(accountScope.retailer);
  const chosenAccounts=()=>masterAccounts.filter(account=>checkedAccounts.has(account.id));
  const exportAccounts=()=>checkedAccounts.size?chosenAccounts():visibleAccounts();
  function resetTransfer(){
    ++credentialView;credentialExportText='';inv('copy-credentials').hidden=true;inv('credentials').value='';inv('transfer-panel').hidden=true;inv('transfer').setAttribute('aria-expanded','false');
  }
  function closeAccountEditor(){
    editedAccount=0;inv('account-form').hidden=true;inv('password').value='';delete inv('password').dataset.edited;++credentialView;
  }
  function selectAccount(id){
    const account=masterAccounts.find(item=>item.id===Number(id));
    editedAccount=account?.id||0;++credentialView;inv('account-form').hidden=false;
    inv('account-heading').textContent=account?'Edit '+retailerTitle(account.retailer)+' account':'New '+selectedRetailer()+' account';
    inv('email').value=account?.email||'';inv('username').value=account?.username||'';
    inv('username').closest('label').hidden=!retailerShowsUsername(account?.retailer??selectedRetailer());
    inv('password').type=ecPrivacyEnabled()?'password':'text';
    inv('password').value='';delete inv('password').dataset.edited;
    inv('password').placeholder=account?.password||account?.hasPassword?'Saved password':'No password set';
    fillDirectoryEditor('accounts');accountEditorChanged();root.dispatchEvent(new Event('ec:account-editor'));
  }
  function paintAccountSelection(){
    for(const row of inv('account-list').querySelectorAll('tr[data-email-summary],tr[data-account-record]')){
      const checkbox=row.cells[0].querySelector('input');
      if(row.hasAttribute('data-email-summary')){
        const accounts=emailAccountGroups.get(row.dataset.emailSummary);
        const count=accounts.reduce((n,account)=>n+Number(checkedAccounts.has(account.id)),0);
        checkbox.checked=count===accounts.length;
        checkbox.indeterminate=count>0&&!checkbox.checked;
      }else checkbox.checked=checkedAccounts.has(Number(row.dataset.accountRecord));
      row.setAttribute('aria-selected',String(checkbox.checked));
    }
  }
  function drawAccountSelection(visible=directoryLogins){
    const count=checkedAccounts.size,all=inv('check-all');
    all.checked=!!visible.length&&count===visible.length;all.indeterminate=count>0&&count<visible.length;all.disabled=!visible.length;
    inv('account-bulk').hidden=!count;inv('selection-count').textContent=count+' selected';closeDirectoryRemoval(inv('delete-confirm'));
    paintAccountSelection();
    inv('transfer').textContent='Export '+(count?count+' selected':visible.length+' visible')+'…';inv('transfer').disabled=!(count||visible.length);
    const records=chosenAccounts(),groups=accountGroups.filter(group=>records.length&&records.every(account=>account.retailer===group.retailer));
    const prior=inv('bulk-group').value;inv('bulk-group').innerHTML=groups.map(g=>'<option value="'+esc(g.id)+'">'+esc(g.name)+'</option>').join('');
    if(groups.some(g=>g.id===prior))inv('bulk-group').value=prior;
    inv('add-group').disabled=!groups.length;inv('remove-group').disabled=!groups.length;
  }
  const treeMarkup=new WeakMap();
  function setTreeMarkup(container,html){
    // Keep nodes in place between the first and second click on a group.
    if(treeMarkup.get(container)===html)return;
    container.innerHTML=html;treeMarkup.set(container,html);
  }
  function enableGroupRename(tree,selector,findGroup,maxLength,renameGroup){
    function start(button){
      const group=findGroup(button);if(!group||button.hidden)return;
      const input=document.createElement('input');
      input.type='text';input.className='ec-group-rename';input.setAttribute('aria-label','Rename group');
      input.maxLength=maxLength;input.autocomplete='off';input.value=group.name;
      button.before(input);button.hidden=true;
      let finished=false;
      function finish(cancel,focus){
        if(finished)return;finished=true;
        const name=input.value.trim();
        const error=cancel||name===group.name?'':!name?'Enter a group name.':renameGroup(group,name);
        if(error&&focus){
          finished=false;input.setCustomValidity(error);input.reportValidity();return;
        }
        input.remove();button.hidden=false;
        if(focus)[...tree.querySelectorAll(selector)].find(node=>findGroup(node)===group)?.focus();
      }
      input.addEventListener('input',()=>input.setCustomValidity(''));
      input.addEventListener('blur',()=>finish(false,false));
      input.addEventListener('keydown',event=>{
        if(event.isComposing||!['Enter','Escape'].includes(event.key))return;
        event.preventDefault();event.stopPropagation();finish(event.key==='Escape',true);
      });
      input.focus();input.select();
    }
    tree.addEventListener('dblclick',event=>{const button=event.target.closest(selector);if(button)start(button);});
    tree.addEventListener('keydown',event=>{
      const button=event.target.closest(selector);
      if(event.key==='F2'&&button){event.preventDefault();start(button);}
    });
  }
  function accountGroupMembers(accounts){
    const groups=new Map(accountGroups.map(group=>[group.id,{retailer:group.retailer,exclusive:[],shared:[]}]));
    for(const account of accounts){
      if(!account.groupIds?.length)continue;
      const ids=[...new Set(account.groupIds)].filter(id=>groups.has(id)&&groups.get(id).retailer===account.retailer);
      for(const id of ids)groups.get(id)[ids.length>1?'shared':'exclusive'].push(account);
    }
    return groups;
  }
  function groupRemovalCounts(members){
    const logins=count=>count+(count===1?' login':' logins');
    return 'Delete '+logins(members.exclusive.length)+' exclusive to this group. Keep '+logins(members.shared.length)+' in other groups.';
  }
  function drawAccountTree(){
    const counts=new Map(),accounts=masterAccounts.filter(managedRetailAccount),members=accountGroupMembers(accounts);
    for(const account of accounts)counts.set(account.retailer,(counts.get(account.retailer)||0)+1);
    const palette=['card-0','card-1','card-2','ec-veronica','ec-account-pool','ec-home'];
    const treeButton=(attr,label,count)=>'<button type="button" '+attr+'>'+esc(label)+' <small>'+count+'</small></button>';
    let html=treeButton('data-all-accounts','All accounts',accounts.length);
    for(const retailer of [...retailerNames(),...(counts.has('')?['']:[])]){
      html+=treeButton('data-retail-node="'+esc(retailer)+'"',retailerTitle(retailer),counts.get(retailer)||0);
      for(const group of accountGroups.filter(g=>g.retailer===retailer).sort((a,b)=>nameCollator.compare(a.name,b.name))){
        const membership=members.get(group.id),note=groupRemovalCounts(membership);
        const count=membership.exclusive.length+membership.shared.length;
        const tint=Array.from(group.id).reduce((hash,letter)=>(hash*31+letter.charCodeAt(0))>>>0,0)%palette.length;
        html+='<div class="ec-chip ec-account-group-chip '+palette[tint]+'"><button type="button" class="ec-tree-link" data-account-group-node="'+esc(group.id)+'" title="Double-click or press F2 to rename"><span>'+esc(group.name)+'</span><small>'+count+'</small></button>'+
          deleteControl('stored-account-group',group.id,group.name+'. '+note)+'</div><p class="ec-account-group-warning" role="status">'+esc(note)+'</p>';
      }
    }
    const tree=inv('account-tree');setTreeMarkup(tree,html);
    for(const button of tree.querySelectorAll('[data-all-accounts],[data-retail-node],[data-account-group-node]')){
      const active=button.hasAttribute('data-all-accounts')?accountScope.retailer===null:
        button.hasAttribute('data-retail-node')?accountScope.retailer===button.dataset.retailNode&&!accountScope.group:accountScope.group===button.dataset.accountGroupNode;
      button.setAttribute('aria-current',String(active));
    }
    inv('group-context').textContent='New group in '+selectedRetailer();
  }
  function accountEmailCell(email,content,retailers,expanded=false){
    const url=parcelMailSearchUrl(email,retailers[0]);
    if(!url)return '<td>'+content+'</td>';
    const multiple=retailers.length>1;
    const label=esc(multiple?'Choose a retailer to view mail for '+email:'View '+(retailers[0]?retailers[0]+' ':'')+'mail for '+email);
    const common=' class="ec-mail-icon" aria-label="'+label+'" title="'+label+'"';
    const icon=multiple?'<button type="button" data-parcel-mail-choose aria-expanded="'+expanded+'"'+common+'>✉︎</button>':
      '<a data-parcel-mail href="'+esc(url)+'" target="_blank" rel="noopener noreferrer" draggable="false"'+common+'>✉︎</a>';
    return '<td><div class="ec-account-email-cell">'+content+icon+'</div></td>';
  }
  function accountLoginRow(account,detail=false){
    const email=esc(account.email||(retailerShowsUsername(account.retailer)?'(username only)':account.username)||'—'),retailer=esc(retailerTitle(account.retailer));
    const editButton='<button type="button" title="Open account details · double-click or press F2 to edit login · drag to a group" data-edit-account="'+account.id+'">'+(detail?retailer:email)+'</button>';
    const selectLabel=esc((detail?retailerTitle(account.retailer)+' ':'')+(account.email||account.username));
    return '<tr draggable="true" data-account-record="'+account.id+'"'+(detail?' class="ec-account-detail"':'')+'><td><input type="checkbox" data-check-account="'+account.id+'" aria-label="Select '+selectLabel+'"></td>'+
      accountEmailCell(account.email,detail?'':editButton,[account.retailer])+'<td>'+(detail?editButton:retailer)+'</td>'+
      (accountUsernamesVisible?'<td>'+esc(retailerShowsUsername(account.retailer)?account.username||'—':'—')+'</td>':'')+
      secretCell(account.id,'password',account.hasPassword||account.password?'••••••':'—')+'</tr>';
  }
  function groupDirectoryLogins(visible){
    emailAccountGroups.clear();
    if(accountScope.retailer!==null)return visible.map(account=>({account}));
    for(const account of visible){
      // A username-only login has its own record even when another retailer uses the same username.
      const key=normalized(account.email)||account.id;
      if(!emailAccountGroups.has(key))emailAccountGroups.set(key,[]);
      emailAccountGroups.get(key).push(account);
    }
    const items=[];
    for(const [key,accounts] of emailAccountGroups){
      if(accounts.length===1){items.push({account:accounts[0]});continue;}
      accounts.sort((a,b)=>nameCollator.compare(a.retailer||'',b.retailer||'')||a.id-b.id);
      items.push({key,accounts});
      if(expandedAccountEmails.has(key))for(const account of accounts)items.push({account,detail:true});
    }
    return items;
  }
  function directoryRow({account,detail,key,accounts}){
    if(account)return accountLoginRow(account,detail);
    const email=esc(accounts[0].email),attr=esc(key),expanded=expandedAccountEmails.has(key);
    const retailers=[...new Set(accounts.map(account=>account.retailer||''))];
    const toggle='<button type="button" data-toggle-email="'+attr+'" aria-label="'+email+'" aria-expanded="'+expanded+'" title="Show or hide retailer logins">'+email+'</button>';
    return '<tr data-email-summary="'+attr+'"><td><input type="checkbox" data-check-email="'+attr+'" aria-label="Select all '+accounts.length+' logins for '+email+'"></td>'+
      accountEmailCell(accounts[0].email,toggle,retailers,expanded)+'<td>'+esc(retailers.map(retailerTitle).join(', '))+'</td>'+
      '<td colspan="'+(accountUsernamesVisible?2:1)+'" class="ec-account-login-count">'+accounts.length+' logins</td></tr>';
  }
  function directoryWindow(body,kind,getItems,renderRow,paintSelection){
    const viewport=body.closest('.ec-record-scroll'),table=body.closest('table');
    let start=0,end=0,frame=0;
    function render(reset=false){
      const items=getItems(),total=items.length,windowed=total>100;
      if(reset){
        body.toggleAttribute('data-windowed',windowed);
        table.setAttribute('aria-rowcount',String(total+1));
      }
      if(!windowed){
        if(reset)body.innerHTML=items.map(renderRow).join('');
        return reset;
      }
      const height=parseFloat(getComputedStyle(table).getPropertyValue('--ec-directory-height'));
      const first=Math.min(total-1,Math.floor(Math.max(0,viewport.scrollTop)/height));
      const visibleEnd=Math.min(total,first+Math.ceil(viewport.clientHeight/height)+1);
      if(!reset&&first>=start+(start?3:0)&&visibleEnd<=end-(end===total?0:3))return false;
      start=Math.max(0,first-10);end=Math.min(total,visibleEnd+10);
      const indices=Array.from({length:end-start},(_,i)=>start+i);
      const focus=!reset&&body.contains(document.activeElement)?document.activeElement.closest('tr[data-directory-index]'):null;
      const pinned=focus?Number(focus.dataset.directoryIndex):-1;
      if(pinned>=0&&pinned<start)indices.unshift(pinned);
      if(pinned>=end)indices.push(pinned);
      const mounted=new Map();
      if(!reset)for(const row of body.querySelectorAll('tr[data-directory-index]'))mounted.set(Number(row.dataset.directoryIndex),row);
      const missing=indices.filter(i=>!mounted.has(i)),batch=document.createElement('tbody');
      batch.innerHTML=missing.map(i=>renderRow(items[i])).join('');
      missing.forEach((i,n)=>{
        const row=batch.rows[n];row.dataset.directoryIndex=i;row.setAttribute('aria-rowindex',String(i+2));mounted.set(i,row);
      });
      const targets=[],columns=Array.from(table.tHead.rows[0].cells).filter(cell=>!cell.hidden).length;
      let previous=-1;
      for(const i of indices){
        if(i>previous+1)targets.push(tableSpacer(i-previous-1,height,columns));
        targets.push(mounted.get(i));previous=i;
      }
      if(previous<total-1)targets.push(tableSpacer(total-previous-1,height,columns));
      updateTableRows(body,targets,focus);return true;
    }
    function schedule(){
      if(frame||getItems().length<=100||sourceTab!==kind||get('ec-sources').hidden)return;
      frame=requestAnimationFrame(()=>{
        frame=0;
        if(sourceTab!==kind||get('ec-sources').hidden||!render())return;
        paintSelection();refreshDirectorySecrets(kind);
      });
    }
    viewport.addEventListener('scroll',schedule,{passive:true});
    if(typeof ResizeObserver==='function')new ResizeObserver(schedule).observe(viewport);
    return {viewport,render};
  }
  let directoryLogins=[],directoryItems=[];
  const directoryAccountsById=new Map();
  const {viewport:loginViewport,render:renderAccountWindow}=directoryWindow(inv('account-list'),'accounts',()=>directoryItems,directoryRow,paintAccountSelection);
  function drawAccounts(preferred){
    retailOptions();drawAccountTree();resetTransfer();
    let visible=visibleAccounts();
    accountUsernamesVisible=accountScope.retailer===null?visible.some(account=>retailerShowsUsername(account.retailer)):retailerShowsUsername(accountScope.retailer);
    root.querySelector('[data-account-sort="username"]').closest('th').hidden=!accountUsernamesVisible;
    if(!accountUsernamesVisible&&accountSort==='username'){accountSort='email';accountDirection=1;visible=visibleAccounts();}
    const ids=new Set(visible.map(a=>a.id));
    for(const id of checkedAccounts)if(!ids.has(id))checkedAccounts.delete(id);
    directoryAccountsById.clear();for(const account of visible)directoryAccountsById.set(account.id,account);
    directoryLogins=visible;directoryItems=groupDirectoryLogins(visible);renderAccountWindow(true);
    for(const button of root.querySelectorAll('[data-account-sort]'))button.closest('th').setAttribute('aria-sort',button.dataset.accountSort===accountSort?(accountDirection===1?'ascending':'descending'):'none');
    inv('account-count').textContent=scopeLabel()+' · '+visible.length.toLocaleString()+(visible.length===1?' login':' logins');
    inv('import-accounts').textContent='Import '+(selectedRetailer()||'retail')+' accounts';drawAccountSelection(visible);
    if(preferred&&ids.has(preferred))selectAccount(preferred);
    else if(editedAccount&&!ids.has(editedAccount))closeAccountEditor();
    refreshDirectorySecrets('accounts');
  }
  function chooseAccountScope(retailer,group=null){
    accountScope={retailer,group};checkedAccounts.clear();closeAccountEditor();
    loginViewport.scrollTop=0;
    if(retailer)inv('retailer').value=retailer;
    inv('search').value='';inv('account-msg').textContent='';drawAccounts();
  }
  inv('account-tree').addEventListener('click',event=>{
    const button=event.target.closest('button');if(!button)return;
    if(button.hasAttribute('data-all-accounts'))chooseAccountScope(null);
    else if(button.hasAttribute('data-retail-node'))chooseAccountScope(button.dataset.retailNode);
    else{const group=accountGroups.find(g=>g.id===button.dataset.accountGroupNode);if(group)chooseAccountScope(group.retailer,group.id);}
  });
  inv('retailer').addEventListener('change',()=>chooseAccountScope(selectedRetailer()));
  inv('search').addEventListener('input',()=>{checkedAccounts.clear();loginViewport.scrollTop=0;drawAccounts();});
  inv('account-list').addEventListener('click',event=>{
    const chooser=event.target.closest('[data-parcel-mail-choose]');
    if(chooser){
      const key=chooser.closest('tr').dataset.emailSummary;
      if(!expandedAccountEmails.has(key)){expandedAccountEmails.add(key);drawAccounts();}
      const row=inv('account-list').querySelector('[data-email-summary="'+CSS.escape(key)+'"]');
      row?.nextElementSibling?.querySelector('[data-parcel-mail]')?.focus();
      return;
    }
    const toggle=event.target.closest('[data-toggle-email]');
    if(toggle){
      const key=toggle.dataset.toggleEmail;
      if(expandedAccountEmails.has(key))expandedAccountEmails.delete(key);else expandedAccountEmails.add(key);
      drawAccounts();
      inv('account-list').querySelector('[data-toggle-email="'+CSS.escape(key)+'"]')?.focus({preventScroll:true});
      return;
    }
    const button=event.target.closest('[data-edit-account]');if(button)selectAccount(button.dataset.editAccount);
  });
  function toggleDirectoryRow(event){
    if(event.detail>1||event.target.closest('input,button,a,label,select,textarea')||!window.getSelection()?.isCollapsed)return;
    const checkbox=event.target.closest('tr[draggable=true],tr[data-email-summary]')?.querySelector('input[type=checkbox]');
    if(checkbox)checkbox.click();
  }
  let textSelectionRow=null;
  function releaseTextSelectionRow(){
    if(textSelectionRow)textSelectionRow.draggable=true;
    textSelectionRow=null;
  }
  for(const body of [inv('account-list'),inv('card-list')]){
    body.addEventListener('click',toggleDirectoryRow);
    // Let the native text selection gesture take priority over group dragging in credential cells.
    body.addEventListener('pointerdown',event=>{
      releaseTextSelectionRow();
      const row=event.target.closest('[data-secret-cell],[data-parcel-mail]')?.closest('tr[draggable=true]');
      if(row){textSelectionRow=row;row.draggable=false;}
    });
  }
  document.addEventListener('pointerup',releaseTextSelectionRow);
  document.addEventListener('pointercancel',releaseTextSelectionRow);
  window.addEventListener('blur',releaseTextSelectionRow);
  function editRetailLogin(event){
    const button=event.target.closest('[data-edit-account]');if(!button)return;
    selectAccount(button.dataset.editAccount);
    const field=inv(inv('email').value||inv('username').closest('label').hidden?'email':'username');
    field.focus();field.select();
  }
  inv('account-list').addEventListener('dblclick',editRetailLogin);
  inv('account-list').addEventListener('keydown',event=>{if(event.key==='F2'){event.preventDefault();editRetailLogin(event);}});
  inv('account-list').addEventListener('change',event=>{
    const box=event.target;
    if(box.hasAttribute('data-check-email')){
      for(const account of emailAccountGroups.get(box.dataset.checkEmail)){
        if(box.checked)checkedAccounts.add(account.id);else checkedAccounts.delete(account.id);
      }
    }else if(box.hasAttribute('data-check-account')){
      const id=Number(box.dataset.checkAccount);
      if(box.checked)checkedAccounts.add(id);else checkedAccounts.delete(id);
    }else return;
    resetTransfer();drawAccountSelection();
  });
  inv('check-all').addEventListener('change',()=>{
    const selected=inv('check-all').checked;checkedAccounts.clear();
    const visible=visibleAccounts();
    if(selected)for(const account of visible)checkedAccounts.add(account.id);
    resetTransfer();drawAccountSelection(visible);
  });
  for(const button of root.querySelectorAll('[data-account-sort]'))button.addEventListener('click',()=>{
    const field=button.dataset.accountSort;accountDirection=field===accountSort?-accountDirection:1;accountSort=field;drawAccounts();
  });
  inv('new-account').addEventListener('click',()=>{selectAccount(0);inv('account-msg').textContent='';inv('email').focus();});
  inv('close-account').addEventListener('click',closeAccountEditor);
  function addAccountToGroup(account,id){
    if(id&&!(account.groupIds||[]).includes(id))account.groupIds=[...(account.groupIds||[]),id];
  }
  function inventoryAccountsChanged(){
    const byId=new Map(masterAccounts.map(a=>[a.id,a]));
    for(const event of events)for(const list of event.lists){
      if(list.rows.some(row=>!byId.has(row.accountId)||!accountMatchesPool(byId.get(row.accountId),list)))rebuildList(event,list);
    }
    refreshSourceRows();renderConfig(false);saveCoordinatorChange({immediate:true});
  }
  inv('group-form').addEventListener('submit',event=>{
    event.preventDefault();const name=inv('account-group-name').value.trim(),retailer=selectedRetailer();
    if(!name||!retailer)return;
    if(accountGroups.some(g=>g.retailer===retailer&&normalized(g.name)===normalized(name))){inv('account-msg').textContent='A group with that name already exists for '+retailer+'.';return;}
    const group={id:'ag-'+crypto.randomUUID(),retailer,name};accountGroups.push(group);
    inv('account-group-name').value='';drawAccounts();inv('bulk-group').value=group.id;
    saveCoordinatorChange({immediate:true});inv('account-msg').textContent='Created '+name+' in '+retailer+'.';
  });
  function renameAccountGroup(group,name){
    if(accountGroups.some(g=>g!==group&&g.retailer===group.retailer&&normalized(g.name)===normalized(name)))return 'That group name is already in use.';
    group.name=name;drawAccounts();renderConfig(false);saveCoordinatorChange({immediate:true});return '';
  }
  enableGroupRename(inv('account-tree'),'[data-account-group-node]',button=>accountGroups.find(g=>g.id===button.dataset.accountGroupNode),100,renameAccountGroup);
  inv('account-tree').addEventListener('ec:confirmed-remove',event=>{
    const {kind,id}=event.detail;
    if(kind!=='stored-account-group')return;
    const index=accountGroups.findIndex(group=>group.id===id);
    if(index<0)return;
    const group=accountGroups[index],members=accountGroupMembers(masterAccounts.filter(managedRetailAccount)).get(id);
    removeRetailAccounts(members.exclusive);
    for(const account of members.shared)account.groupIds=account.groupIds.filter(key=>key!==id);
    accountGroups.splice(index,1);
    for(const item of events)for(const list of item.lists)if(list.accountGroupIds?.includes(id))list.accountGroupIds=list.accountGroupIds.filter(key=>key!==id);
    if(accountScope.group===id)chooseAccountScope(group.retailer);else drawAccounts();
    inventoryAccountsChanged();
    inv('account-msg').textContent='Removed '+group.name+'. Deleted '+members.exclusive.length+' exclusive and kept '+members.shared.length+' shared logins.';
  });
  function changeAccountMembership(group,ids,add){
    const accounts=masterAccounts.filter(account=>ids.has(account.id));
    if(!group||!accounts.length||accounts.some(account=>account.retailer!==group.retailer))return;
    const changed=accounts.filter(account=>Boolean(account.groupIds?.includes(group.id))!==add);
    for(const account of changed){
      if(add)addAccountToGroup(account,group.id);
      else account.groupIds=account.groupIds.filter(id=>id!==group.id);
    }
    if(changed.length){drawAccounts();inventoryAccountsChanged();}
    inv('account-msg').textContent=changed.length?(add?'Added to ':'Removed from ')+group.name+' · '+changed.length+(changed.length===1?' account.':' accounts.'):
      (add?'Already in ':'No selected accounts in ')+group.name+'.';
  }
  for(const add of [true,false])inv(add?'add-group':'remove-group').addEventListener('click',()=>{
    changeAccountMembership(accountGroups.find(group=>group.id===inv('bulk-group').value),checkedAccounts,add);
  });
  inv('assign-retailer').addEventListener('click',()=>{
    const retailer=inv('bulk-retailer').value,selected=chosenAccounts(),seen=new Set();
    if(!retailer||!selected.length)return;
    for(const a of masterAccounts.filter(a=>managedRetailAccount(a)&&a.retailer===retailer)){
      for(const key of [a.email,a.username].filter(Boolean).map(normalized))seen.add(key);
    }
    for(const account of selected.filter(a=>a.retailer!==retailer)){
      const keys=new Set([account.email,account.username].filter(Boolean).map(normalized));
      if([...keys].some(k=>seen.has(k))){inv('account-msg').textContent='Assignment would duplicate a login at '+retailer+'. No accounts changed.';return;}
      for(const key of keys)seen.add(key);
    }
    for(const account of selected)if(account.retailer!==retailer){account.retailer=retailer;account.groupIds=[];}
    chooseAccountScope(retailer);inventoryAccountsChanged();inv('account-msg').textContent='Assigned '+selected.length+' accounts to '+retailer+'.';
  });
  function removeRetailAccounts(records){
    if(!records.length)return;
    const ids=new Set(records.map(account=>account.id));
    for(const account of records)if(account.parcelKey)removedAccountKeys.add(account.parcelKey);
    forgetAccountSecrets(records);
    for(let i=masterAccounts.length-1;i>=0;i--)if(ids.has(masterAccounts[i].id))masterAccounts.splice(i,1);
  }
  inv('delete-accounts').addEventListener('click',()=>{
    if(!checkedAccounts.size)return;
    inv('delete-note').textContent='Delete '+checkedAccounts.size+' selected accounts and their credentials from Event Coordinator?';
    openDirectoryRemoval(inv('delete-accounts'),inv('delete-confirm'),inv('delete-cancel'));
  });
  inv('delete-cancel').addEventListener('click',()=>closeDirectoryRemoval(inv('delete-confirm')));
  inv('delete-apply').addEventListener('click',()=>{
    if(!confirmDirectoryRemoval(inv('delete-confirm'))||!checkedAccounts.size)return;
    const removed=chosenAccounts();removeRetailAccounts(removed);
    checkedAccounts.clear();closeAccountEditor();drawAccounts();inventoryAccountsChanged();inv('account-msg').textContent='Deleted '+removed.length+' accounts.';
  });
  inv('account-form').addEventListener('submit',event=>{
    event.preventDefault();const email=inv('email').value.trim(),username=inv('username').value.trim(),password=inv('password').value;
    let account=masterAccounts.find(item=>item.id===editedAccount);const retailer=account?account.retailer:selectedRetailer();
    if(!email&&!username){inv('account-msg').textContent=retailerShowsUsername(retailer)?'Enter an email or username.':'Enter an email.';return;}
    const logins=new Set([email,username].filter(Boolean).map(normalized));
    const matching=masterAccounts.filter(a=>a.id!==editedAccount&&a.retailer===retailer&&[a.email,a.username].some(value=>value&&logins.has(normalized(value))));
    if(matching.some(managedRetailAccount)){
      inv('account-msg').textContent='This retailer already has that login.';return;
    }
    if(!account){
      account=matching.find(a=>email&&normalized(a.email)===normalized(email));
      if(!account){account={id:nextAccountId(),retailer,password:''};masterAccounts.push(account);}
      addAccountToGroup(account,accountScope.group);
    }
    if(account.parcelKey)account.retailManaged=true;
    Object.assign(account,{email,username});if(password&&inv('password').dataset.edited){account.password=password;invalidateDirectorySecrets();}
    inv('search').value='';drawAccounts(account.id);inventoryAccountsChanged();inv('account-msg').textContent='Saved '+retailerTitle(retailer)+' account.';
  });
  inv('transfer').addEventListener('click',()=>{
    if(!inv('transfer-panel').hidden){resetTransfer();return;}
    resetTransfer();inv('transfer-panel').hidden=false;inv('transfer').setAttribute('aria-expanded','true');
    inv('transfer-heading').textContent='Export '+exportAccounts().length+(checkedAccounts.size?' selected':' visible')+' accounts · '+scopeLabel();
    inv('account-msg').textContent='';inv('export').focus();
  });
  inv('import-accounts').addEventListener('click',()=>{resetTransfer();openMappedImport('accounts');});
  inv('import-cards').addEventListener('click',()=>openMappedImport('cards'));
  inv('hide').addEventListener('click',resetTransfer);
  inv('credentials').addEventListener('copy',event=>{
    const field=inv('credentials');
    if(ecPrivacyEnabled()&&credentialExportText&&field.selectionStart===0&&field.selectionEnd===field.value.length){
      event.clipboardData.setData('text/plain',credentialExportText);event.preventDefault();
    }
  });
  inv('copy-credentials').addEventListener('click',async()=>{
    if(!credentialExportText)return;
    try{await navigator.clipboard.writeText(credentialExportText);inv('account-msg').textContent='Copied credentials.';}
    catch{inv('credentials').select();inv('account-msg').textContent='Press Copy to copy the selected credentials.';}
  });
  inv('export').addEventListener('click',async()=>{
    const token=++credentialView,eligible=exportAccounts().filter(a=>a.hasPassword||a.password);
    try{
      const passwords=await readRetailPasswords(eligible);
      if(token!==credentialView||sourceTab!=='accounts'||get('ec-sources').hidden||inv('transfer-panel').hidden)return;
      const rows=eligible.filter(account=>passwords[account.id]);
      credentialExportText=rows.map(account=>(account.email||account.username)+':'+passwords[account.id]).join('\n');
      inv('copy-credentials').hidden=!rows.length;inv('credentials').value=rows.map(account=>(account.email||account.username)+':'+ecPrivateText(passwords[account.id])).join('\n');
      inv('credentials').select();inv('account-msg').textContent=rows.length+' credentials ready to copy.';
    }catch{inv('account-msg').textContent='Could not open saved credentials.';}
  });
  drawAccounts();
  const accountImportFields=[
    {key:'email',label:'Email',aliases:['email address','account email','retailer email','e-mail']},
    {key:'username',label:'Username',aliases:['user','user name','handle','login','login name','account login']},
    {key:'password',label:'Password',aliases:['pass','passwd','pwd','account password'],sensitive:true}
  ];
  function inspectAccountImport(rows,context){
    const retailer=selectedRetailer(),issues=[],plans=[],claims=new Map(),accountsByLogin=new Map();
    let skipped=0,added=0,updated=0;
    if(!retailer)issues.push('Choose a retailer in the account directory before importing.');
    if(!context.mapping.some(key=>key==='email'||key==='username'))issues.push('Map an email or username column.');
    if(!issues.length){
      for(const account of selectedAccounts()){
        for(const login of new Set([account.email,account.username].filter(Boolean).map(normalized))){
          if(!accountsByLogin.has(login))accountsByLogin.set(login,new Set());
          accountsByLogin.get(login).add(account);
        }
      }
      for(const row of rows){
        const email=(row.email||'').trim(),username=(row.username||'').trim(),password=row.password||'';
        const problem=message=>issues.push('Row '+row.__row+' · '+message);
        if(!email&&!username){problem('Enter an email or username.');continue;}
        if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){problem('Check the mapped email address.');continue;}
        if(password.length>4096){problem('Password exceeds 4,096 characters.');continue;}
        const logins=[...new Set([email,username].filter(Boolean).map(normalized))];
        const matches=[...new Set(logins.flatMap(login=>[...(accountsByLogin.get(login)||[])]))];
        const existing=matches.filter(managedRetailAccount);
        const candidates=existing.length?existing:matches.filter(item=>email&&normalized(item.email)===normalized(email));
        const [account]=candidates;
        if(candidates.length>1||account?.email&&email&&normalized(account.email)!==normalized(email)){
          problem('The mapped logins identify different accounts. Review the email and username.');continue;
        }
        const keys=logins.map(login=>'login:'+login);
        if(account)keys.push('id:'+account.id);
        const signature=JSON.stringify([normalized(email),normalized(username),password]);
        const prior=keys.map(key=>claims.get(key)).filter(Boolean);
        if(prior.length){
          if(prior.every(entry=>entry.signature===signature))skipped++;
          else problem('Conflicts with row '+prior[0].line+' for this retailer.');
          continue;
        }
        for(const key of keys)claims.set(key,{line:row.__row,signature});
        if(account&&managedRetailAccount(account)&&(!accountScope.group||account.groupIds?.includes(accountScope.group))&&(!email||normalized(account.email)===normalized(email))&&(!username||account.username===username)&&(!password||account.password===password)){skipped++;continue;}
        plans.push({account,email,username,password});
        if(account&&managedRetailAccount(account))updated++;else added++;
      }
    }
    return {
      summary:added+' new · '+updated+' updated · '+skipped+' skipped',issues,actions:plans.length,
      notes:['Matching logins update this retailer’s account. Empty fields keep saved values.'],
      async apply(){
        if(issues.length||!plans.length)return;
        if(retailer!==selectedRetailer())throw Error('The selected retailer changed. Reopen the import.');
        let first;
        for(const plan of plans){
          const account=plan.account||{id:nextAccountId(),retailer,email:'',username:'',password:''};
          if(!plan.account)masterAccounts.push(account);
          if(account.parcelKey)account.retailManaged=true;
          addAccountToGroup(account,accountScope.group);
          for(const key of ['email','username','password'])if(plan[key])account[key]=plan[key];
          first??=account.id;
        }
        invalidateDirectorySecrets();resetTransfer();inv('search').value='';drawAccounts(first);
        refreshSourceRows();renderConfig(false);await finishInventoryImport('accounts');
        inv('account-msg').textContent=retailer+' accounts imported · '+added+' new · '+updated+' updated.';
      }
    };
  }
  mappedImportAdapters.accounts={
    fields:accountImportFields,title:()=>'Import '+(selectedRetailer()||'retail')+' accounts'+(accountScope.group?' · '+groupTitle(accountScope.group):''),
    preset:{separator:':',mapping:['email','password'],label:'email:password'},inspect:inspectAccountImport
  };
  const cardImportFields=[
    {key:'label',label:'Card label',aliases:['label','nickname','card nickname','card alias','alias']},
    {key:'groups',label:'Card groups',aliases:['group','groups','card group','group membership']},
    {key:'lastFour',label:'Last four digits',aliases:['last four','last4','last 4','ending']},
    {key:'number',label:'Card number',aliases:['number','cc number','ccnumber','pan','credit card number','payment card number'],sensitive:true},
    {key:'cvv',label:'Security code (CVV / CVC)',aliases:['CVV2','CVC','CVC2','CCV','CID','CSC','security code','card security code','payment security code','card verification code','card verification value','card CVV','card CVC','cc CVV','cc CVC'],sensitive:true},
    {key:'holder',label:'Cardholder name',aliases:['holder','name on card','card name','cardholder','card holder','card name optional']},
    {key:'expiry',label:'Expiry (MM/YY or MM/YYYY)',aliases:['expiry','expiration','exp','exp date','expiry date','expiration date']},
    {key:'month',label:'Expiry month',aliases:['month','expiration month','expire month','exp month','cc exp month','cc month']},
    {key:'year',label:'Expiry year',aliases:['year','expiration year','expire year','exp year','cc exp year','cc year']},
    {key:'brand',label:'Card brand',aliases:['brand','card type','network']},
    {key:'paymentType',label:'Payment method',aliases:['payment type','payment method optional']},
    {key:'firstName',label:'Billing first name',aliases:['billing first','bill first name']},
    {key:'lastName',label:'Billing last name',aliases:['billing last','bill last name']},
    {key:'street',label:'Billing street',aliases:['billing address','billing address line 1','billing street 1']},
    {key:'street2',label:'Billing street 2',aliases:['billing address 2','billing address line 2']},
    {key:'city',label:'Billing city',aliases:['billing town']},
    {key:'state',label:'Billing state',aliases:['billing region','billing province']},
    {key:'zip',label:'Billing postal code',aliases:['billing zip','billing zip code','billing post code','billing postcode','billing postcode / zip']},
    {key:'country',label:'Billing country',aliases:['billing country code']}
  ];
  const billingKeys=['firstName','lastName','street','street2','city','state','zip','country'];
  function readImportedCard(row){
    const clean=key=>(row[key]||'').trim();
    const details={};
    const code=row.cvv||'';
    if(code){
      if(code.length<3||code.length>4||/[^0-9]/.test(code))throw Error('Use a 3- or 4-digit security code (CVV / CVC).');
      details.cvv=code;
    }
    for(const key of ['number','month','year','holder','brand','paymentType',...billingKeys]){
      const value=clean(key);
      if(value)details[key]=value;
    }
    if(Object.values(details).some(value=>value.length>500))throw Error('A checkout field exceeds 500 characters.');
    if(details.number){
      details.number=details.number.replace(/[\s-]/g,'');
      const issue=cardNumberIssue(details.number);
      if(issue)throw Error(issue);
    }
    const expiry=clean('expiry');
    if(expiry){
      const parts=expiry.match(/^(\d{1,2})\s*[/-]\s*(\d{4}|\d{2})$/);
      if(!parts)throw Error('Use MM/YY or MM/YYYY for the expiry date.');
      if(details.month&&details.month.padStart(2,'0')!==parts[1].padStart(2,'0')||details.year&&details.year.slice(-2)!==parts[2].slice(-2))throw Error('The combined expiry conflicts with the separate month or year.');
      details.month??=parts[1];details.year??=parts[2];
    }
    if(details.month){if(!/^(0?[1-9]|1[0-2])$/.test(details.month))throw Error('Use an expiry month from 01 to 12.');details.month=details.month.padStart(2,'0');}
    if(details.year){if(!/^(?:\d{2}|\d{4})$/.test(details.year))throw Error('Use a two- or four-digit expiry year.');if(details.year.length===2)details.year='20'+details.year;}
    const lastFour=clean('lastFour')||details.number?.slice(-4)||'';
    if(lastFour&&(!/^\d{4}$/.test(lastFour)||details.number&&!details.number.endsWith(lastFour)))throw Error('Check the last four digits against the card number.');
    const label=clean('label')||(lastFour?'Card ending '+lastFour:'');
    if(!label||label.length>80)throw Error('Supply a card label up to 80 characters, a card number, or the last four digits.');
    const groups=clean('groups').split(/[;|]/).map(value=>value.trim()).filter(Boolean);
    if(groups.some(value=>value.length>80))throw Error('Group names can contain up to 80 characters.');
    return {label,lastFour,groups,details};
  }
  function inspectCardImport(rows,context){
    const issues=[],plans=[],seen=new Map(),existing=new Map(masterCards.map(label=>[normalized(label),label]));
    let skipped=0,updated=0;
    if(!context.mapping.some(key=>['label','number','lastFour'].includes(key)))issues.push('Map a card label, card number, or last four digits.');
    else for(const row of rows){
      try{
        const card=readImportedCard(row);
        if(!row.label?.trim()&&card.lastFour){
          const legacy=existing.get(normalized('Card •••• '+card.lastFour));
          if(legacy){
            if(existing.has(normalized(card.label)))throw Error('Two saved labels use this ending. Map the exact card label.');
            card.label=legacy;
          }
        }
        const key=normalized(card.label),signature=JSON.stringify(card);
        if(seen.has(key)){
          const prior=seen.get(key);
          if(prior.signature!==signature)throw Error('This card label conflicts with row '+prior.line+'. Give different cards distinct labels.');
          skipped++;continue;
        }
        seen.set(key,{signature,line:row.__row});
        if(existing.has(key)){
          const label=existing.get(key),savedLastFour=cardLastFour.get(label);
          if(savedLastFour&&card.lastFour&&savedLastFour!==card.lastFour)throw Error('The card ending does not match the saved label.');
          if(!card.details.cvv){skipped++;continue;}
          if(!row.label?.trim()&&!card.details.number)throw Error('Map the saved card label or full card number to update its security code.');
          card.label=label;card.existingId=cardIds.get(label);card.line=row.__row;updated++;
        }
        plans.push(card);
      }catch(error){issues.push('Row '+row.__row+' · '+error.message);}
    }
    return {
      summary:(plans.length-updated)+' new · '+updated+' updated · '+skipped+' skipped',issues,actions:plans.length,
      notes:['A matching label updates its security code when supplied. Other saved details and groups are kept. Separate new card groups with |.'],
      async apply(){
        if(issues.length||!plans.length)return;
        const withNumber=plans.filter(card=>card.existingId&&card.details.number);
        if(withNumber.length){
          const saved=await resolveCheckout(withNumber.map(card=>card.existingId));
          for(const card of withNumber){
            const number=saved.get(card.existingId)?.number;
            if(number&&number!==card.details.number)throw Error('Row '+card.line+' · The card number does not match the saved label. Review the label before importing.');
          }
        }
        const recipes=captureCardRecipes();
        const groupsByName=new Map(cardGroups.map(group=>[normalized(group.name),group]));
        for(const {label,lastFour,groups,details,existingId} of plans){
          if(existingId){stageImportedCard(existingId,{cvv:details.cvv});continue;}
          const id='card-'+nextCardId++;masterCards.push(label);cardIds.set(label,id);cardLastFour.set(label,lastFour);
          const network=displayNetwork(details.brand,details.number);
          if(network)cardNetworks.set(id,network);
          if(Object.keys(details).length)stageImportedCard(id,details);
          if(billingKeys.some(key=>details[key]))cardBilling.set(id,Object.fromEntries(billingKeys.map(key=>[key,details[key]||''])));
          for(const name of [...groups,...(cardScope?[cardScope]:[])]){
            const key=normalized(name);
            if(!groupsByName.has(key)){const group={name,cards:[]};cardGroups.push(group);groupsByName.set(key,group);}
            const members=groupsByName.get(key).cards;if(!members.includes(label))members.push(label);
          }
        }
        refreshCardRecipes(recipes);drawCardGroups();drawCards(plans[0].label);refreshSourceRows();renderConfig(false);await finishInventoryImport('cards');
        inv('card-msg').textContent=plans.length+(plans.length===1?' card':' cards')+' imported.';
      }
    };
  }
  mappedImportAdapters.cards={fields:cardImportFields,title:()=>cardScope?'Import cards into '+cardScope:'Import cards',inspect:inspectCardImport};
  let cardScope=null,cardSort='label',cardOrder=1,cardEditorOpen=false;
  const checkedCards=new Set();
  let directoryCards=[];
  function cardDirectoryRow({id,label,network,lastFour}){
    return '<tr draggable="true" data-card-record="'+esc(id)+'" aria-selected="'+checkedCards.has(id)+'"><td><input type="checkbox" data-check-card="'+esc(id)+'" aria-label="Select '+esc(label)+'" '+(checkedCards.has(id)?'checked':'')+'></td><td><button type="button" title="Open card details · drag to a group" data-open-card="'+esc(label)+'">'+esc(label)+'</button></td><td>'+esc(network||'—')+'</td>'+secretCell(id,'number',lastFour?'•••• '+lastFour:'—',lastFour?'Ending '+lastFour:'Not set')+secretCell(id,'cvv',cardsWithSecurityCode.has(id)?'••••':'—')+'</tr>';
  }
  const {viewport:cardViewport,render:renderCardWindow}=directoryWindow(inv('card-list'),'cards',()=>directoryCards,cardDirectoryRow,paintCardSelection);
  let forgetCardSecrets=id=>{sampleCardDetails.delete(id);};
  function visibleCardRecords(){
    const query=normalized(get('ec-card-search').value);
    const records=[];
    for(const label of masterCards){
      const id=cardIds.get(label),names=cardGroupNames(label);
      if(cardScope!==null&&(cardScope?!names.includes(cardScope):names.length>0))continue;
      const record={id,label,groups:names.join(', '),network:cardNetworks.get(id)||'',lastFour:cardLastFour.get(label)||''};
      if(!query||[label,record.groups,record.network,record.lastFour].some(value=>normalized(value).includes(query)))records.push(record);
    }
    return records.sort((a,b)=>cardOrder*nameCollator.compare(a[cardSort],b[cardSort])||nameCollator.compare(a.label,b.label));
  }
  function drawCardTree(groups){
    const tree=get('ec-card-tree'),rows=get('ec-card-group-rows');
    const html=groups.map(group=>{
      const name=esc(group.name);
      return '<div class="ec-card-group-row" data-card-group-row="'+name+'"><button type="button" class="ec-tree-link" data-card-scope="group" data-group-card="'+name+'" title="Double-click or press F2 to rename"><span>'+name+'</span><small>'+group.cards.length+'</small></button>'+deleteControl('stored-card-group',group.name,group.name)+'</div>';
    }).join('');
    setTreeMarkup(rows,html);
    tree.querySelector('[data-card-scope=all] small').textContent=masterCards.length;
    tree.querySelector('[data-card-scope=ungrouped] small').textContent=masterCards.filter(card=>!cardGroupNames(card).length).length;
    for(const button of tree.querySelectorAll('[data-card-scope]')){
      const scope=button.dataset.cardScope==='all'?null:button.dataset.groupCard||'';
      button.setAttribute('aria-current',String(scope===cardScope));
    }
  }
  function drawCardDirectory(){
    if(cardScope&&!cardGroups.some(group=>group.name===cardScope))cardScope=null;
    const groups=[...cardGroups].sort((a,b)=>nameCollator.compare(a.name,b.name));
    drawCardTree(groups);
    get('ec-card-scope').textContent=cardScope===null?'All cards':cardScope||'Ungrouped';
    const bulkGroup=get('ec-card-bulk-group'),previous=bulkGroup.value;
    bulkGroup.innerHTML=groups.map(group=>'<option>'+esc(group.name)+'</option>').join('');
    if(groups.some(group=>group.name===previous))bulkGroup.value=previous;
    else if(cardScope)bulkGroup.value=cardScope;
    const records=visibleCardRecords(),visible=new Set(records.map(record=>record.id));
    directoryCards=records;
    for(const id of checkedCards)if(!visible.has(id))checkedCards.delete(id);
    renderCardWindow(true);
    get('ec-card-result-count').textContent=records.length.toLocaleString()+(records.length===1?' card':' cards');
    for(const button of get('ec-settings-cards').querySelectorAll('[data-card-sort]')){
      button.parentElement.setAttribute('aria-sort',cardSort===button.dataset.cardSort?(cardOrder===1?'ascending':'descending'):'none');
    }
    drawCardSelection();refreshDirectorySecrets('cards');
  }
  function paintCardSelection(){
    for(const row of inv('card-list').querySelectorAll('tr[data-card-record]')){
      const selected=checkedCards.has(row.dataset.cardRecord);
      row.setAttribute('aria-selected',String(selected));row.querySelector('[data-check-card]').checked=selected;
    }
  }
  function drawCardSelection(){
    const all=get('ec-all-cards'),count=checkedCards.size,total=directoryCards.length;
    all.disabled=!total;all.checked=!!count&&count===total;all.indeterminate=count>0&&count<total;
    paintCardSelection();
    get('ec-card-bulk').hidden=!count;get('ec-card-selection-count').textContent=count.toLocaleString()+' selected';
    get('ec-card-membership').hidden=!count||!cardGroups.length;
    get('ec-card-membership-count').textContent=count.toLocaleString()+(count===1?' selected card':' selected cards');
    closeDirectoryRemoval(get('ec-card-delete-confirm'));
  }
  function closeCardEditor(){
    for(const field of [inv('card-number'),inv('card-code')])field.value='';
    root.dispatchEvent(new CustomEvent('ec:card-editor'));cardEditorOpen=false;editedCard='';inv('card-form').hidden=true;
  }
  function drawCardEditor(label){
    inv('card-form').hidden=!cardEditorOpen;
    root.dispatchEvent(new CustomEvent('ec:card-editor'));
    editedCard=masterCards.includes(label)?label:'';
    for(const field of [inv('card-number'),inv('card-code')]){field.value='';field.type=ecPrivacyEnabled()?'password':'text';delete field.dataset.edited;field.setCustomValidity('');}
    inv('card-code').placeholder=cardsWithSecurityCode.has(cardIds.get(editedCard))?'Saved · enter to replace':'3–4 digits';
    fillDirectoryEditor('cards');inv('card-name').value=editedCard;inv('four').value=cardLastFour.get(editedCard)||'';
    inv('card-heading').textContent=editedCard?'Edit '+editedCard:'New card';root.dispatchEvent(new Event('ec:card-editor-ready'));
  }
  function drawCards(label=editedCard){
    drawCardEditor(label);drawCardDirectory();
  }
  get('ec-card-tree').addEventListener('click',event=>{
    const button=event.target.closest('[data-card-scope]');if(!button)return;
    cardScope=button.dataset.cardScope==='all'?null:button.dataset.groupCard||'';
    get('ec-card-search').value='';cardViewport.scrollTop=0;checkedCards.clear();closeCardEditor();drawCardDirectory();
    if(cardScope)get('ec-card-bulk-group').value=cardScope;
  });
  get('ec-card-search').addEventListener('input',()=>{checkedCards.clear();cardViewport.scrollTop=0;drawCardDirectory();});
  get('ec-settings-cards').addEventListener('click',event=>{
    const sort=event.target.closest('[data-card-sort]');if(!sort)return;
    cardOrder=cardSort===sort.dataset.cardSort?-cardOrder:1;cardSort=sort.dataset.cardSort;cardViewport.scrollTop=0;drawCardDirectory();
  });
  inv('card-list').addEventListener('click',event=>{
    const button=event.target.closest('[data-open-card]');if(!button)return;
    cardEditorOpen=true;drawCardEditor(button.dataset.openCard);refreshDirectorySecrets('cards');inv('card-name').focus();
  });
  inv('card-list').addEventListener('change',event=>{
    const box=event.target.closest('[data-check-card]');if(!box)return;
    if(box.checked)checkedCards.add(box.dataset.checkCard);else checkedCards.delete(box.dataset.checkCard);
    drawCardSelection();
  });
  get('ec-all-cards').addEventListener('change',event=>{
    checkedCards.clear();if(event.target.checked)for(const record of directoryCards)checkedCards.add(record.id);
    drawCardSelection();
  });
  function setCardGroupForm(open){
    const form=get('ec-card-group-create'),input=get('ec-card-directory-group-name'),toggle=get('ec-add-card-group');
    form.hidden=!open;toggle.setAttribute('aria-expanded',String(open));input.setCustomValidity('');
    if(open)input.focus();else{input.value='';toggle.focus();}
  }
  get('ec-add-card-group').addEventListener('click',()=>setCardGroupForm(get('ec-card-group-create').hidden));
  get('ec-card-directory-group-name').addEventListener('input',event=>event.target.setCustomValidity(''));
  get('ec-card-group-create').addEventListener('keydown',event=>{
    if(event.key==='Escape'){event.preventDefault();setCardGroupForm(false);}
  });
  get('ec-card-group-create').addEventListener('submit',event=>{
    event.preventDefault();
    const input=get('ec-card-directory-group-name'),name=input.value.trim();
    if(!name||cardGroups.some(item=>normalized(item.name)===normalized(name))){
      input.setCustomValidity(!name?'Enter a group name.':'That group name is already in use.');input.reportValidity();return;
    }
    cardGroups.push({name,cards:[]});
    cardScope=name;checkedCards.clear();get('ec-card-search').value='';
    setCardGroupForm(false);closeCardEditor();drawCardGroups();drawCardDirectory();get('ec-card-bulk-group').value=name;
    saveCoordinatorChange({immediate:true});
    get('ec-card-tree').querySelector('[data-group-card][aria-current=true]')?.focus();
  });
  function changeStoredCardGroup(group,newName){
    if(newName!==null&&cardGroups.some(item=>item!==group&&normalized(item.name)===normalized(newName)))return 'That group name is already in use.';
    const recipes=captureCardRecipes();
    if(cardScope===group.name)cardScope=newName;
    if(newName===null)cardGroups.splice(cardGroups.indexOf(group),1);else group.name=newName;
    refreshCardRecipes(recipes);
    drawCardGroups();drawCardDirectory();renderConfig(false);saveCoordinatorChange({immediate:true});return '';
  }
  enableGroupRename(get('ec-card-tree'),'[data-group-card]',button=>cardGroups.find(g=>g.name===button.dataset.groupCard),80,changeStoredCardGroup);
  get('ec-card-tree').addEventListener('ec:confirmed-remove',event=>{
    if(event.detail?.kind!=='stored-card-group')return;
    const group=cardGroups.find(g=>g.name===event.detail.id);if(!group)return;
    changeStoredCardGroup(group,null);
    inv('card-msg').textContent='Removed '+group.name+'. Its cards remain available.';
    get('ec-card-tree').querySelector('[data-card-scope][aria-current=true]')?.focus();
  });
  function changeCardMembership(group,ids,adding){
    if(!group||!ids.size)return;
    const membership=new Set(group.cards);
    const changed=masterCards.filter(label=>ids.has(cardIds.get(label))&&membership.has(label)!==adding);
    if(changed.length){
      const recipes=captureCardRecipes();
      for(const label of changed){if(adding)membership.add(label);else membership.delete(label);}
      group.cards=[...membership];
      refreshCardRecipes(recipes);drawCardGroups();
      if(cardEditorOpen&&editedCard&&cardScope!==null&&!visibleCardRecords().some(card=>card.label===editedCard))closeCardEditor();
      drawCardDirectory();renderConfig(false);saveCoordinatorChange({immediate:true});
    }
    inv('card-msg').textContent=changed.length?(adding?'Added to ':'Removed from ')+group.name+' · '+changed.length+(changed.length===1?' card.':' cards.'):
      (adding?'Already in ':'No selected cards in ')+group.name+'.';
  }
  for(const adding of [true,false])get(adding?'ec-card-bulk-add':'ec-card-bulk-remove').addEventListener('click',()=>{
    changeCardMembership(cardGroups.find(group=>group.name===get('ec-card-bulk-group').value),checkedCards,adding);
  });
  get('ec-card-delete').addEventListener('click',()=>{
    if(!checkedCards.size)return;
    get('ec-card-confirm-delete').textContent='Confirm delete '+checkedCards.size;
    openDirectoryRemoval(get('ec-card-delete'),get('ec-card-delete-confirm'),get('ec-card-cancel-delete'));
  });
  get('ec-card-confirm-delete').addEventListener('click',()=>{
    if(!confirmDirectoryRemoval(get('ec-card-delete-confirm'))||!checkedCards.size)return;
    const recipes=captureCardRecipes(),labels=new Set(masterCards.filter(label=>checkedCards.has(cardIds.get(label))));
    invalidateDirectorySecrets();
    for(const label of labels){
      const id=cardIds.get(label);cardIds.delete(label);cardLastFour.delete(label);cardNetworks.delete(id);cardBilling.delete(id);cardsWithSecurityCode.delete(id);forgetCardSecrets(id);
    }
    masterCards.splice(0,masterCards.length,...masterCards.filter(label=>!labels.has(label)));
    for(const group of cardGroups)group.cards=group.cards.filter(label=>!labels.has(label));
    refreshCardRecipes(recipes);
    inv('card-msg').textContent=labels.size+(labels.size===1?' card deleted.':' cards deleted.');
    checkedCards.clear();closeCardEditor();drawCardGroups();drawCardDirectory();renderConfig(false);saveCoordinatorChange({immediate:true});
  });
  function enableGroupDrops({body,tree,checked,recordId,groupAttribute,eligibleGroups,assign,noun}){
    let drag=null;
    const selector='['+groupAttribute+']';
    function finish(){
      if(!drag)return;
      drag.preview.remove();
      for(const target of tree.querySelectorAll('.ec-drop-ready,.ec-drop-over'))target.classList.remove('ec-drop-ready','ec-drop-over');
      drag=null;
    }
    body.addEventListener('dragstart',event=>{
      if(event.target.closest('[data-parcel-mail]')){event.preventDefault();return;}
      const row=event.target.closest('tr[draggable=true]');
      if(!row||!event.dataTransfer)return;
      finish();
      const id=recordId(row),ids=checked.has(id)?new Set(checked):new Set([id]);
      const groups=eligibleGroups(ids),preview=document.createElement('span');
      if(!groups.size){event.preventDefault();return;}
      preview.className='ec-drag-preview';preview.textContent=ids.size+' '+noun+(ids.size===1?'':'s');root.append(preview);
      drag={ids,groups,preview};
      event.dataTransfer.effectAllowed='copy';
      event.dataTransfer.setData('application/x-event-coordinator-group','records');
      event.dataTransfer.setDragImage(preview,0,12);
      for(const target of tree.querySelectorAll(selector)){
        if(groups.has(target.getAttribute(groupAttribute)))target.classList.add('ec-drop-ready');
      }
    });
    function dropTarget(element){
      const button=element.closest(selector)||element.closest('[data-card-group-row]')?.querySelector(selector);
      return button&&tree.contains(button)&&drag?.groups.has(button.getAttribute(groupAttribute))?button:null;
    }
    tree.addEventListener('dragover',event=>{
      if(!drag)return;
      const target=dropTarget(event.target);
      for(const button of tree.querySelectorAll('.ec-drop-over'))if(button!==target)button.classList.remove('ec-drop-over');
      if(target){event.preventDefault();event.dataTransfer.dropEffect='copy';target.classList.add('ec-drop-over');}
    });
    tree.addEventListener('dragleave',event=>{
      if(drag&&!tree.contains(event.relatedTarget))tree.querySelector('.ec-drop-over')?.classList.remove('ec-drop-over');
    });
    tree.addEventListener('drop',event=>{
      const target=dropTarget(event.target);
      if(!target)return;
      event.preventDefault();
      const {ids}=drag,key=target.getAttribute(groupAttribute);
      finish();assign(key,ids);
    });
    document.addEventListener('dragend',finish);
    window.addEventListener('pagehide',finish);
  }
  enableGroupDrops({body:inv('card-list'),tree:get('ec-card-tree'),checked:checkedCards,recordId:row=>row.dataset.cardRecord,
    groupAttribute:'data-group-card',eligibleGroups:()=>new Set(cardGroups.map(group=>group.name)),
    assign:(name,ids)=>changeCardMembership(cardGroups.find(group=>group.name===name),ids,true),noun:'card'});
  enableGroupDrops({body:inv('account-list'),tree:inv('account-tree'),checked:checkedAccounts,recordId:row=>Number(row.dataset.accountRecord),
    groupAttribute:'data-account-group-node',eligibleGroups:ids=>{
      const retailers=new Set(masterAccounts.filter(account=>ids.has(account.id)).map(account=>account.retailer));
      return new Set(accountGroups.filter(group=>retailers.size===1&&retailers.has(group.retailer)).map(group=>group.id));
    },assign:(id,ids)=>changeAccountMembership(accountGroups.find(group=>group.id===id),ids,true),noun:'account'});
  get('ec-card-cancel-delete').addEventListener('click',()=>closeDirectoryRemoval(get('ec-card-delete-confirm')));
  inv('new-card').addEventListener('click',()=>{cardEditorOpen=true;drawCardEditor('');refreshDirectorySecrets('cards');inv('card-msg').textContent='';inv('card-name').focus();});
  inv('card-form').addEventListener('submit',event=>{
    event.preventDefault();const label=inv('card-name').value.trim(),original=editedCard;
    if(!label||masterCards.some(card=>card!==original&&card.toLowerCase()===label.toLowerCase())){inv('card-msg').textContent='Use a unique card label.';return;}
    const numberInput=inv('card-number'),codeInput=inv('card-code');
    const number=numberInput.value.replace(/[\s-]/g,'');
    const recipes=captureCardRecipes();
    if(original){masterCards[masterCards.indexOf(original)]=label;cardLastFour.delete(original);const identity=cardIds.get(original);cardIds.delete(original);cardIds.set(label,identity);}else{masterCards.push(label);cardIds.set(label,'card-'+nextCardId++);}
    const id=cardIds.get(label),patch={};
    if(numberInput.dataset.edited&&number){patch.number=number;inv('four').value=number.slice(-4);const network=detectedNetwork(number);if(network)cardNetworks.set(id,network);else cardNetworks.delete(id);}
    if(codeInput.dataset.edited&&codeInput.value)patch.cvv=codeInput.value;
    if(Object.keys(patch).length)stageImportedCard(id,patch);
    cardLastFour.set(label,inv('four').value);
    if(!original&&cardScope)cardGroups.find(group=>group.name===cardScope)?.cards.push(label);
    if(original&&original!==label){
      for(const group of cardGroups)group.cards=group.cards.map(card=>card===original?label:card);
      for(const eventConfig of events)for(const list of eventConfig.lists){
        if(list.cards.delete(original))list.cards.add(label);for(const row of list.rows)if(row.card===original)row.card=label;
      }
    }
    refreshCardRecipes(recipes,original,label);drawCardGroups();renderConfig(false);drawCards(label);inv('card-msg').textContent='Saved card.';
  });
  get('ec-list-retailer').addEventListener('change',()=>{const list=currentList();list.retailer=get('ec-list-retailer').value;list.accountGroupIds=null;refreshLive();});
  get('ec-account-group-options').addEventListener('change',event=>{
    const box=event.target,list=currentList();if(!list)return;
    if(box.hasAttribute('data-pool-all'))list.accountGroupIds=box.checked?null:[];
    else if(box.dataset.poolGroup){const chosen=new Set(list.accountGroupIds||[]);if(box.checked)chosen.add(box.dataset.poolGroup);else chosen.delete(box.dataset.poolGroup);list.accountGroupIds=[...chosen];}
    else return;
    list.pool=true;get('ec-pool').checked=true;refreshLive();saveCoordinatorChange({immediate:true});
  });
  root.addEventListener('ec:privacy-change',()=>{
    invalidateDirectorySecrets();resetTransfer();
    if(get('ec-sources').hidden)return;
    if(sourceTab==='accounts'||sourceTab==='cards')refreshDirectorySecrets(sourceTab);
  });
  root.addEventListener('ec:show-source',event=>{resetTransfer();closeAccountEditor();if(event.detail==='accounts')drawAccounts();if(event.detail==='cards')drawCards(editedCard||masterCards[0]);});
  get('ec-show-coordinator').addEventListener('click',()=>{resetTransfer();inv('password').value='';});
  const checkoutLayouts={
    sniped:[
      ['PROFILE GROUP (optional)','group'],['PROFILE NAME (optional)','nickname'],['FIRST NAME','firstName'],['LAST NAME','lastName'],['EMAIL','email'],['PHONE NUMBER','phone'],['ADDRESS LINE 1','street'],['ADDRESS LINE 2','street2'],['CITY','city'],['STATE','state'],['POSTCODE / ZIP','zip'],['COUNTRY','country'],
      ['CARD NUMBER','number'],['EXPIRE MONTH','month'],['EXPIRE YEAR','year'],['CARD CVC','cvv'],['BILLING ADDRESS LINE 1','billingStreet'],['BILLING ADDRESS LINE 2','billingStreet2'],['BILLING CITY','billingCity'],['BILLING STATE','billingState'],['BILLING POSTCODE / ZIP','billingZip'],['BILLING COUNTRY','billingCountry'],['CARD NAME (optional)','holder'],['PAYMENT METHOD (optional)','paymentType']],
    shikari:[
      ['profile_name','nickname'],['first_name','firstName'],['last_name','lastName'],['email','email'],['phone_num','phone'],['cc_number','number'],['cc_exp_month','month'],['cc_exp_year','year'],['cc_cvv','cvv'],['shipping_street','street'],['shipping_street_2','street2'],['shipping_city','city'],['shipping_state','state'],['shipping_zip_code','zip'],['shipping_country','country'],['billing_first_name','billingFirstName'],['billing_last_name','billingLastName'],['billing_street','billingStreet'],['billing_street_2','billingStreet2'],['billing_city','billingCity'],['billing_state','billingState'],['billing_zip_code','billingZip'],['billing_country','billingCountry']],
    aycd:[
      ['Email Address','email'],['Profile Name','nickname'],['Only One Checkout','oneCheckout'],['Name on Card','holder'],['Card Type','brand'],['Card Number','number'],['Expiration Month','month'],['Expiration Year','year'],['CVV','cvv'],['Same Billing/Shipping','sameAddress'],['Shipping Name','shippingName'],['Shipping Phone','phone'],['Shipping Address','street'],['Shipping Address 2','street2'],['Shipping Address 3','street3'],['Shipping Post Code','zip'],['Shipping City','city'],['Shipping State','state'],['Shipping Country','country'],['Billing Name','billingName'],['Billing Phone','billingPhone'],['Billing Address','billingStreet'],['Billing Address 2','billingStreet2'],['Billing Address 3','billingStreet3'],['Billing Post Code','billingZip'],['Billing City','billingCity'],['Billing State','billingState'],['Billing Country','billingCountry'],['Size (Optional)','size']]
  };
  const checkoutText=value=>String(value??'');
  function targetAddressText(value){
    return String(value??'').replace(/\./g,'').replace(/\s*,\s*/g,' ').replace(/[ \t]{2,}/g,' ').trim();
  }
  function formatTargetAddress(record,street,second,third){
    const seen=new Set(),extra=[];
    for(const field of [street,second,third]){
      const main=field===street;
      let text=String(record[field]||'');
      if(!main&&/^\s*\d+[a-z]?(?:[-/][\da-z]+)?\s*$/i.test(text))text='#'+text.trim();
      // Preserve non-address instructions, including gate codes and apartment references in prose.
      const pattern=/(?:(apartment|apt|suite|ste|unit|building|bldg|floor|fl|room|rm)(?:\.\s*#?\s*|\s*#\s*|\s+|(?=\d))|#\s*)([a-z]+[-/]?\d[\w/-]*|\d[\w/-]*|[a-z])(?=$|[\s,;:.()])/gi;
      let end=0,moved=false;
      let line=text.replace(pattern,(whole,label,id,index)=>{
        if(index&&/[\w#]/.test(text[index-1]))return whole;
        const prefix=text.slice(end,index);
        if(!main&&prefix.trim()&&!/[,;\n]\s*$/.test(prefix))return whole;
        const kind=/^(building|bldg)$/i.test(label)?'building':/^(floor|fl)$/i.test(label)?'floor':/^(room|rm)$/i.test(label)?'room':'unit';
        const key=kind+' '+id.toLowerCase();
        const unit=whole.replace(/^([a-z]+)\.(?=[a-z\d])/i,'$1 ');
        end=index+whole.length;moved=true;
        if(seen.has(key))return '';
        seen.add(key);
        if(main)return unit;
        extra.push(unit);return '';
      });
      if(!main&&moved)line=line.replace(/^[\s,;:-]+|[\s,;:-]+$/g,'');
      record[field]=targetAddressText(line);
    }
    record[street]=[record[street],...extra.map(targetAddressText)].filter(Boolean).join(' ');
  }
  function formatTargetExport(record){
    formatTargetAddress(record,'street','street2','street3');
    formatTargetAddress(record,'billingStreet','billingStreet2','billingStreet3');
    for(const field of ['city','state','country','billingCity','billingState','billingCountry'])record[field]=targetAddressText(record[field]);
    for(const field of ['zip','billingZip'])record[field]=String(record[field]??'').replace(/[.,]/g,'');
  }
  function exportedNickname(nickname,format){
    const name=checkoutText(nickname).trim(),prefix={sniped:'Sniped',shikari:'Shikari'}[format];
    if(!name||!prefix||name.toLowerCase().startsWith(prefix.toLowerCase()+' '))return name;
    return prefix+' '+name;
  }
  const paymentFields=['holder','brand','number','month','year','cvv','paymentType'];
  const exportFlag=value=>value===true||String(value).trim().toLowerCase()==='true';
  function profilePostal(value,country){
    const postal=String(value??'').trim();
    const domestic=/^(US|USA|United States(?: of America)?)$/i.test(String(country||'US').replace(/\./g,'').trim());
    const extended=domestic&&/^(\d{5})[- ]?(\d{4})$/.exec(postal);
    return extended?extended[1]:postal;
  }
  function checkoutRecord(row,card,cvv,format,disabled=new Set(),retailer=''){
    const info=card||{},number=checkoutText(info.number).replace(/[ -]/g,''),month=checkoutText(info.month).trim(),year=checkoutText(info.year).trim();
    const record={...row,nickname:exportedNickname(row.nickname,format),number,month:/^\d{1,2}$/.test(month)?month.padStart(2,'0'):month,year:/^\d{2}$/.test(year)?'20'+year:year,
      cvv,holder:info.holder||'',brand:info.brand||'',paymentType:info.paymentType||'',
      billingPhone:row.billingPhone??row.phone};
    if(retailer==='Target')formatTargetExport(record);
    record.zip=profilePostal(record.zip,record.country);
    record.billingZip=profilePostal(record.billingZip,record.billingCountry||record.country);
    for(const key of disabled)record[key]='';
    const accountEmail=checkoutText(row.accountEmail).trim();
    record.email=!disabled.has('email')&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(accountEmail)?accountEmail:'';
    record.billingEmail=record.email;
    if(disabled.has('card'))for(const key of paymentFields)record[key]='';
    const fullName=(first,last,original)=>disabled.has(first)||disabled.has(last)?[record[first],record[last]].filter(Boolean).join(' '):original??[record[first],record[last]].filter(Boolean).join(' ');
    record.shippingName=fullName('firstName','lastName',row.shippingName);
    record.billingName=fullName('billingFirstName','billingLastName',row.billingName);
    record.oneCheckout=String(exportFlag(record.onlyCheckoutOnce));
    record.matchNameOnCardAndAddress=!disabled.has('card')&&!disabled.has('holder')&&exportFlag(record.matchNameOnCardAndAddress);
    const shipping=aycdAddress(record,false),billing=aycdAddress(record,true);
    record.sameAddress=String(!disabled.has('sameBillingAndShippingAddress')&&Object.keys(shipping).every(key=>shipping[key]===billing[key]));
    return record;
  }
  const regionNames={
    US:{AL:'Alabama',AK:'Alaska',AZ:'Arizona',AR:'Arkansas',CA:'California',CO:'Colorado',CT:'Connecticut',DE:'Delaware',DC:'District of Columbia',FL:'Florida',GA:'Georgia',HI:'Hawaii',ID:'Idaho',IL:'Illinois',IN:'Indiana',IA:'Iowa',KS:'Kansas',KY:'Kentucky',LA:'Louisiana',ME:'Maine',MD:'Maryland',MA:'Massachusetts',MI:'Michigan',MN:'Minnesota',MS:'Mississippi',MO:'Missouri',MT:'Montana',NE:'Nebraska',NV:'Nevada',NH:'New Hampshire',NJ:'New Jersey',NM:'New Mexico',NY:'New York',NC:'North Carolina',ND:'North Dakota',OH:'Ohio',OK:'Oklahoma',OR:'Oregon',PA:'Pennsylvania',RI:'Rhode Island',SC:'South Carolina',SD:'South Dakota',TN:'Tennessee',TX:'Texas',UT:'Utah',VT:'Vermont',VA:'Virginia',WA:'Washington',WV:'West Virginia',WI:'Wisconsin',WY:'Wyoming',PR:'Puerto Rico',GU:'Guam',AS:'American Samoa',MP:'Northern Mariana Islands',VI:'U.S. Virgin Islands'},
    CA:{AB:'Alberta',BC:'British Columbia',MB:'Manitoba',NB:'New Brunswick',NL:'Newfoundland and Labrador',NS:'Nova Scotia',NT:'Northwest Territories',NU:'Nunavut',ON:'Ontario',PE:'Prince Edward Island',QC:'Quebec',SK:'Saskatchewan',YT:'Yukon'}
  };
  const countryNames=typeof Intl.DisplayNames==='function'?new Intl.DisplayNames(['en'],{type:'region'}):null;
  function aycdAddress(record,billing){
    const value=(shippingKey,billingKey)=>checkoutText(record[billing?billingKey:shippingKey]);
    const country=value('country','billingCountry'),state=value('state','billingState'),upper=country.toUpperCase();
    const code=({'USA':'US','UNITED STATES':'US','UNITED STATES OF AMERICA':'US','CAN':'CA','CANADA':'CA','UK':'GB'})[upper]||upper;
    const readable=/^[A-Z]{2}$/.test(code)?countryNames?.of(code):null;
    return {name:value('shippingName','billingName'),email:value('email','billingEmail'),phone:value('phone','billingPhone'),
      line1:value('street','billingStreet'),line2:value('street2','billingStreet2'),line3:value('street3','billingStreet3'),
      postCode:value('zip','billingZip'),city:value('city','billingCity'),country:readable||country,state:regionNames[code]?.[state.toUpperCase()]||state};
  }
  function encodeAycdJSON(records){
    const brands={mastercard:'MasterCard',visa:'Visa',amex:'American Express',americanexpress:'American Express',discover:'Discover'};
    return JSON.stringify(records.map(record=>({
      name:checkoutText(record.nickname),notes:checkoutText(record.notes),billingAddress:aycdAddress(record,true),shippingAddress:aycdAddress(record,false),
      paymentDetails:{nameOnCard:checkoutText(record.holder),cardType:brands[checkoutText(record.brand).toLowerCase().replace(/[^a-z]/g,'')]||checkoutText(record.brand),
        cardNumber:checkoutText(record.number),cardExpMonth:checkoutText(record.month),cardExpYear:checkoutText(record.year),cardCvv:checkoutText(record.cvv)},
      sameBillingAndShippingAddress:exportFlag(record.sameAddress),onlyCheckoutOnce:exportFlag(record.oneCheckout),matchNameOnCardAndAddress:record.matchNameOnCardAndAddress
    })),null,2)+'\n';
  }
  const quoteCheckout=value=>{const text=checkoutText(value);return /[",\r\n]/.test(text)?'"'+text.replace(/"/g,'""')+'"':text;};
  function checkoutMatrix(format,records){
    const columns=checkoutLayouts[format];
    return [columns.map(([label])=>label),...records.map(record=>columns.map(([,key])=>checkoutText(record[key])))];
  }
  function encodeCheckoutCSV(format,records){
    return checkoutMatrix(format,records).map(row=>row.map(quoteCheckout).join(',')).join('\r\n')+'\r\n';
  }
  function checkoutProblems(format,row,disabled=new Set()){
    const needs=['nickname','phone','street','city','country','billingStreet','billingCity','billingCountry'];
    if(format==='aycd')needs.push('shippingName','billingName');
    else needs.push('firstName','lastName');
    if(format==='shikari')needs.push('billingFirstName','billingLastName');
    if(['US','USA','UNITED STATES','CA','CAN','CANADA'].includes(checkoutText(row.country).toUpperCase()))needs.push('state','zip');
    if(['US','USA','UNITED STATES','CA','CAN','CANADA'].includes(checkoutText(row.billingCountry).toUpperCase()))needs.push('billingState','billingZip');
    const fullNameLabels={shippingName:'Shipping name',billingName:'Billing name'};
    const omitted=key=>disabled.has(key)||(key==='shippingName'&&disabled.has('firstName')&&disabled.has('lastName'))||(key==='billingName'&&disabled.has('billingFirstName')&&disabled.has('billingLastName'));
    const issues=needs.filter(key=>!omitted(key)&&!checkoutText(row[key]).trim()).map(key=>fullNameLabels[key]||profileColumns.find(([field])=>field===key)?.[1]||key);
    if(!disabled.has('email')&&!row.email)issues.push('Retail account email');
    if(disabled.has('card'))return issues;
    if(!disabled.has('number')&&!/^\d{12,19}$/.test(row.number))issues.push('card number');
    if(!disabled.has('month')&&(!/^\d{2}$/.test(row.month)||Number(row.month)<1||Number(row.month)>12))issues.push('expiration month');
    if(!disabled.has('year')&&!/^\d{4}$/.test(row.year))issues.push('expiration year');
    if(!disabled.has('cvv')&&!/^\d{3,4}$/.test(row.cvv))issues.push('security code');
    return issues;
  }
  let checkoutScope=null,checkoutRequest=0,checkoutBusy=false;
  const exportDialog=get('ec-export-dialog'),exportResult=get('ec-export-result');
  function clearCheckout(){
    checkoutRequest++;checkoutScope=null;checkoutBusy=false;get('ec-export-codes').replaceChildren();get('ec-export-issue').replaceChildren();get('ec-download-export').disabled=true;
  }
  function closeCheckout(){clearCheckout();if(exportDialog.open)exportDialog.close();}
  function checkoutProblemList(issues){
    const panel=get('ec-export-issue');panel.replaceChildren();if(!issues.length)return;
    const heading=document.createElement('strong');heading.textContent='Export warnings';panel.append(heading);
    const list=document.createElement('ul');for(const message of issues){const item=document.createElement('li');item.textContent=message;list.append(item);}panel.append(list);
  }
  function setExportFileTypes(){
    const aycd=get('ec-export-format').value==='aycd',types=get('ec-export-filetype'),json=types.querySelector('[value=json]');
    json.disabled=!aycd;json.hidden=!aycd;
    if(aycd)types.value='json';
    else if(!types.value||types.value==='json')types.value='csv';
  }
  function showFormatNote(format,kind,records,disabled){
    const note=get('ec-format-note');
    let message='';
    if(format==='sniped'){
      message='Sniped accepts separate shipping and billing addresses and uses the shipping recipient and phone.';
      const different=records.filter(row=>row.billingName&&row.billingName!==row.shippingName||row.billingPhone&&row.billingPhone!==row.phone).length;
      if(different)message+=' The shipping contact will be used for '+different+' '+(different===1?'profile':'profiles')+' with different billing contact details.';
    }else if(format==='shikari'){
      message='Shikari uses the shipping phone.';
    }else if(format==='aycd'&&kind==='json'){
      message='AYCD native JSON has an email field in each address.';
    }
    const emailNote=disabled.has('email')?'The Email column is disabled. Exported email fields will be blank.':
      'Each exported email field uses the assigned retail account’s email.';
    note.textContent=[message,emailNote].filter(Boolean).join(' ');note.hidden=false;
  }
  function refreshCheckout(){
    get('ec-download-export').disabled=true;if(!checkoutScope)return;
    const {list,cards,rows}=checkoutScope,format=get('ec-export-format').value,warnings=[],codes=new Map([...get('ec-export-codes').querySelectorAll('input')].map(input=>[input.dataset.ref,input.value])),records=[];
    const kind=get('ec-export-filetype').value,disabled=list.disabledColumns||new Set();
    const fileHelp={json:'In AYCD Profiles, import using the AYCD format.',xlsx:'Every cell is stored as text, including card numbers and leading zeros.',
      csv:(format==='aycd'?'In AYCD Profiles, import using “CSV/Google Sheet/Form”. ':'')+'For editing in Excel or Numbers, choose the .xlsx file type to keep every digit.'};
    const targetHint=normalized(list.retailer||'Target')==='target'?' If Target suggests an address at checkout, choose it after checking the apartment or suite.':'';
    get('ec-export-help').textContent=fileHelp[kind]+' This file includes checkout details.'+targetHint;
    const pending=rows.reduce((count,row)=>count+Number(!!row.pending),0);
    if(pending)warnings.push(pending.toLocaleString()+(pending===1?' pending profile is':' pending profiles are')+' included. Missing account or card details will be blank.');
    if(disabledDestinationCount(list,allShippingDestinations()))warnings.push('This export includes a disabled shipping destination.');
    const event=currentEvent();
    if(!event.sharing){const others=otherListAccounts(event,list);if(rows.some(row=>others.has(row.accountId)))warnings.push('Some accounts are also assigned to other lists in this event.');}
    const fields=new Map(),checkedCards=new Set();
    const paymentIssues=new Set(['card number','expiration month','expiration year','security code']);
    for(const row of rows){
      const id=cardIds.get(row.card),card=cards.get(id);
      const rec=checkoutRecord(row,card,codes.get(id)??card?.cvv??'',format,disabled,list.retailer||'Target');
      const issues=checkoutProblems(format,rec,disabled);
      for(const issue of issues){
        const payment=paymentIssues.has(issue);
        if(payment&&(!row.card||checkedCards.has(row.card)))continue;
        let field=fields.get(issue);
        if(!field){field={count:0,sample:payment?row.card:(row.nickname||'Unnamed profile')};fields.set(issue,field);}
        field.count++;
      }
      checkedCards.add(row.card);
      records.push(rec);
    }
    for(const [issue,{count,sample}] of fields){
      const subject=paymentIssues.has(issue)?'card':'profile';
      warnings.push(issue+' missing or invalid on '+count.toLocaleString()+' '+subject+(count===1?'':'s')+' ('+(count===1?'':'e.g. ')+sample+').');
    }
    showFormatNote(format,kind,records,disabled);
    checkoutProblemList(warnings);
    const button=get('ec-download-export');
    button.textContent='Export '+kind.toUpperCase()+(warnings.length?' anyway':'');
    button.disabled=!records.length||checkoutBusy;
    return records.length?records:null;
  }
  function exportBaseName(value,byteLimit=200){
    let text=value.normalize('NFC').replace(/[<>:"/\\|?*\p{Cc}\p{Cf}]/gu,'-').trim().replace(/^\.+|[. ]+$/g,'');
    if(/^(con|prn|aux|nul|com[1-9¹²³]|lpt[1-9¹²³])(?:\.|$)/i.test(text))text='_'+text;
    const encoder=new TextEncoder();
    while(encoder.encode(text).length>byteLimit)text=Array.from(text).slice(0,-1).join('');
    return text.replace(/[. ]+$/g,'')||'Profiles';
  }
  function exportFilename(name,kind,now){
    const two=n=>String(n).padStart(2,'0'),hours=now.getHours();
    const date=two(now.getMonth()+1)+'-'+two(now.getDate())+'-'+now.getFullYear();
    const clock=(hours%12||12)+'-'+two(now.getMinutes())+(hours<12?' AM':' PM');
    const base=exportBaseName(name.trim().replace(/\.(?:csv|json|xlsx)$/i,''),200-date.length-clock.length-2);
    return base+' '+date+' '+clock+'.'+kind;
  }
  async function openCheckout(){
    clearCheckout();exportResult.textContent='';const list=currentList();if(!list)return;
    const request=checkoutRequest,rows=orderedRows(list,true),eventId=activeEvent;
    if(!rows.length)return;
    get('ec-export-omissions').textContent=list.disabledColumns.size?omissionSummary(list.disabledColumns,populatedCounts(rows)):'';
    get('ec-export-format').value=list.format;setExportFileTypes();get('ec-export-name').value=list.name;get('ec-export-context').textContent=currentEvent().name+' / '+list.name+' · '+rows.length.toLocaleString()+' profiles';
    get('ec-format-note').hidden=true;
    get('ec-export-issue').textContent='Checking selected cards…';get('ec-export-help').textContent='';exportDialog.showModal();
    try{
      const used=new Map(rows.map(row=>[cardIds.get(row.card),row.card])),ids=list.disabledColumns?.has('card')?[]:[...used.keys()].filter(Boolean);
      const cards=ids.length&&typeof getExportCardDetails==='function'?await getExportCardDetails(ids):new Map();
      if(request!==checkoutRequest||!exportDialog.open||currentList()!==list)return;
      const needsCodes=list.disabledColumns?.has('cvv')?[]:ids.filter(id=>!/^\d{3,4}$/.test(cards.get(id)?.cvv||''));
      checkoutScope={list,eventId,cards,rows};get('ec-export-codes').innerHTML=needsCodes.map(id=>'<label>'+esc(used.get(id))+' · Security code (CVV/CVC, optional)<input data-ec-private data-ref="'+esc(id)+'" type="'+(ecPrivacyEnabled()?'password':'text')+'" maxlength="4" inputmode="numeric" autocomplete="off" placeholder="3–4 digits"></label>').join('');refreshCheckout();
    }catch(error){if(request===checkoutRequest)get('ec-export-issue').textContent='Could not load cards for this workspace. Close and retry after reconnecting.';}
  }
  get('ec-open-export').addEventListener('click',openCheckout);
  get('ec-export-filetype').addEventListener('change',refreshCheckout);
  get('ec-export-format').addEventListener('change',()=>{setExportFileTypes();refreshCheckout();});get('ec-export-codes').addEventListener('input',refreshCheckout);
  get('ec-close-export').addEventListener('click',closeCheckout);exportDialog.addEventListener('cancel',event=>{event.preventDefault();if(!checkoutBusy)closeCheckout();});exportDialog.addEventListener('close',clearCheckout);
  window.addEventListener('pagehide',closeCheckout);
  root.addEventListener('click',event=>{if(exportDialog.open&&event.target.closest('.ec-worknav button,[data-list],[data-event]'))closeCheckout();},true);
  get('ec-download-export').addEventListener('click',async()=>{
    if(checkoutBusy||!checkoutScope)return;
    if(checkoutScope.list!==currentList()||checkoutScope.eventId!==activeEvent){closeCheckout();return;}
    const records=refreshCheckout();if(!records)return;
    const format=get('ec-export-format').value,list=currentList(),event=currentEvent(),instant=new Date(),id=Array.from(crypto.getRandomValues(new Uint8Array(8)),byte=>byte.toString(16).padStart(2,'0')).join('');
    const kind=get('ec-export-filetype').value;
    const name=get('ec-export-name').value.trim()||list.name;
    let filename=exportFilename(name,kind,instant);
    const accounts=new Map(masterAccounts.map(account=>[account.id,account])),summary={id,eventId:event.id,eventName:event.name,listId:list.id,listName:list.name,bot:format,format,at:instant.toISOString(),profiles:checkoutScope.rows.map((row,index)=>({accountId:accounts.get(row.accountId)?.parcelKey??row.accountId??null,addressId:row.parentAddressId,...(row.addressId!==row.parentAddressId?{destinationId:row.addressId}:{}),cardId:list.disabledColumns?.has('card')?null:cardIds.get(row.card)??null,nickname:records[index].nickname}))};
    checkoutBusy=true;get('ec-download-export').disabled=true;get('ec-close-export').disabled=true;
    try{
      const payload=kind==='xlsx'?makeTextWorkbook(checkoutMatrix(format,records)):kind==='json'?encodeAycdJSON(records):encodeCheckoutCSV(format,records);
      const type={xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',json:'application/json',csv:'text/csv'}[kind];
      if(typeof privateExportFile==='function'){filename=await privateExportFile(filename,payload,kind,type)||filename;}else{
        const url=URL.createObjectURL(new Blob([payload],{type})),anchor=document.createElement('a');anchor.href=url;anchor.download=filename;root.append(anchor);anchor.click();anchor.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);
      }
      list.format=format;
      closeCheckout();exportResult.textContent=records.length+' profiles exported · '+filename;
      if(typeof recordProfileExport==='function'){try{await recordProfileExport(summary);}catch{exportResult.textContent+=' · history could not be saved';}}
    }catch(error){
      if(error?.name==='ExportFileExistsError'){
        checkoutBusy=false;get('ec-download-export').disabled=false;get('ec-export-issue').textContent=error.message;
      }else{closeCheckout();exportResult.textContent=error?.name==='AbortError'?'Export cancelled.':'Export failed. Reopen Export profiles to try again.';}
    }finally{get('ec-close-export').disabled=false;}
  });
  (() => {
  const scope=document.querySelector('meta[name="parcel-workspace"]')?.content||'';
  const header=root.querySelector('.ec-brand .ec-sample'),reload=document.createElement('button'),notice=document.createElement('div');
  reload.type='button';reload.className='cursor-interaction';reload.textContent='Refresh workspace';header.before(reload);
  notice.className='ec-source-note ec-parcel-notice';notice.setAttribute('role','status');root.querySelector('.ec-brand').after(notice);
  let profileScope='',loading=false,accountSerial=0,active=false;
  let revision=0,csrf='',lastAck='',persistTimer=0,saving=null,conflicted=false;
  const pendingCards=new Map(),pendingPasswords=new Map();
  const accountIds=new Map(),lastSources=new Map();
  // Keep shared defaults valid for older clients. Each tab remembers its actual
// selection separately so browsing a different event never changes a revision.
let sharedEventId=null,sharedListIds=new Map(),tabNavigation=null;
const tabNavigationKey=()=>(scope||profileScope)?'event-coordinator.view.'+(scope||profileScope):'';
function clearSharedNavigation(){
  sharedEventId=null;sharedListIds.clear();tabNavigation=null;
}
function rememberSharedNavigation(state){
  sharedEventId=state.selected.event;
  sharedListIds=new Map(state.events.map(event=>[event.id,state.version===1?
    (event.id===sharedEventId?state.selected.bot:'sniped'):event.selectedListId]));
}
function sharedSelectedEvent(){
  return events.some(event=>event.id===sharedEventId)?sharedEventId:events[0]?.id??null;
}
function sharedSelectedList(event){
  const id=sharedListIds.get(event.id);
  return event.lists.some(list=>list.id===id)?id:event.lists[0]?.id??null;
}
rememberTabView=()=>{
  const key=tabNavigationKey();
  if(!key)return;
  tabNavigation={event:activeEvent,lists:events.map(event=>[event.id,event.selectedListId])};
  try{sessionStorage.setItem(key,JSON.stringify(tabNavigation));}catch{}
};
function restoreTabNavigation(){
  if(!tabNavigation){
    const key=tabNavigationKey();if(!key)return;
    try{tabNavigation=JSON.parse(sessionStorage.getItem(key));}catch{}
  }
  if(Array.isArray(tabNavigation?.lists)){
    const lists=new Map(tabNavigation.lists.filter(pair=>Array.isArray(pair)&&pair.length===2));
    for(const event of events){
      const id=lists.get(event.id);
      if(event.lists.some(list=>list.id===id))event.selectedListId=id;
    }
    if(events.some(event=>event.id===tabNavigation.event))activeEvent=tabNavigation.event;
  }
  rememberTabView();
}

// Keep only touched DOM nodes, never copies of the card/password fields.
const draftForms=new WeakSet(),cardDraftInputs=new Set();
let accountDraft=false,cardEditorId=null;
for(const type of ['input','change'])root.addEventListener(type,event=>{
  const input=event.target,form=input.closest('form');
  if(!form)return;
  if(form===inv('card-form'))cardDraftInputs.add(input);
  else if(form===inv('account-form'))accountDraft=true;
  else if(get('ec-source-editor').contains(form)&&!form.dataset.sourceAddress)draftForms.add(form);
});
root.addEventListener('ec:account-editor',()=>{accountDraft=false;});
root.addEventListener('ec:card-editor',()=>{cardDraftInputs.clear();cardEditorId=null;});
root.addEventListener('ec:card-editor-ready',()=>{cardEditorId=cardIds.get(editedCard)||null;});
function cardEditorHasDraft(){
  for(const input of cardDraftInputs){
    if(input===inv('card-code')||input===inv('card-number')){if(input.dataset.edited)return true;}
    else if(!detailsFields.contains(input)||editingDetails)return true;
  }
  return false;
}
function canApplyRemoteView(){
  if(root.querySelector('dialog[open],.ec-strip-rename,.ec-group-rename,.ec-drop-ready')||armedChip||pendingDirectoryRemoval)return false;
  const focus=document.activeElement;
  if(root.contains(focus)&&focus.matches('input:not([type=checkbox]):not([type=button]),textarea,select,[contenteditable=true]'))return false;
  if(get('ec-sources').hidden)return !cellDrag&&!selectingCells&&bulkValue.hidden;
  if(sourceTab==='cards')return (inv('card-form').hidden||(!cardEditorHasDraft()&&!detailsToggle.disabled))
      &&!get('ec-card-directory-group-name').value;
  if(sourceTab==='accounts')return (inv('account-form').hidden||!accountDraft)&&inv('transfer-panel').hidden
      &&!inv('account-group-name').value;
  const editor=get('ec-source-editor'),form=editor.querySelector('form');
  for(const input of editor.querySelectorAll('[data-suite-recipient-input],[name=newRecipient]'))if(input.value)return false;
  if(form?.dataset.sourceAddress)return form.querySelector('[type=submit]').disabled;
  return !draftForms.has(form);
}
function redrawSyncedWorkspace(){
  const scroller=get('ec-sheet-scroll'),top=scroller.scrollTop,left=scroller.scrollLeft;
  drawAddressPicker();updatePalette();retailOptions();
  if(!get('ec-coordinator-view').hidden){
    renderView(false);
    scroller.scrollTop=top;scroller.scrollLeft=left;
  }
  if(get('ec-sources').hidden)return;
  if(sourceTab==='cards'){
    if(!cardEditorOpen){drawCardDirectory();return;}
    const label=masterCards.find(label=>cardIds.get(label)===cardEditorId),expanded=editingDetails;
    if(cardEditorId&&!label){closeCardEditor();drawCardDirectory();return;}
    drawCards(label||'');
    if(expanded)detailsToggle.click();
    return;
  }
  if(sourceTab==='accounts'){
    if(accountScope.group&&!accountGroups.some(group=>group.id===accountScope.group))accountScope.group=null;
    drawAccounts(inv('account-form').hidden?0:editedAccount);return;
  }
  const address=directoryAddresses().find(item=>item.id===selectedAddress);
  if((sourceMode==='address'&&!address)||(sourceMode==='person'&&!ownerOf(selectedPerson))){
    const fallback=directoryAddresses().find(item=>item.personId===selectedPerson)||directoryAddresses()[0];
    selectedAddress=fallback?.id??null;selectedPerson=fallback?.personId??people[0]?.id??null;
    sourceMode=fallback?'address':selectedPerson?'person':'new-person';
  }else if(address&&sourceMode==='address')selectedPerson=address.personId;
  drawDirectory();
}

  function requestError(status,message){return Object.assign(new Error(message),{status});}
  let sessionRequest=null;
  async function privateHeaders(){
    if(!csrf){
      if(!sessionRequest)sessionRequest=(async()=>{
        const response=await fetch('/api/session',{credentials:'same-origin',mode:'same-origin',cache:'no-store',redirect:'error'});
        if(!response.ok)throw requestError(response.status,'Could not check the local session.');
        const result=await response.json();
        if(!result.csrf||!result.user)throw requestError(401,'Reload this workspace.');
        if(result.user.profile!==(scope||profileScope))throw requestError(403,'Reload this workspace.');
        csrf=result.csrf;
      })().finally(()=>{sessionRequest=null;});
      await sessionRequest;
    }
    return {'Content-Type':'application/json','X-Parcel-Workspace':scope||profileScope,'X-Chaos-CSRF':csrf};
  }
  async function privateRequest(path,options){
    const send=headers=>fetch(path,{credentials:'same-origin',mode:'same-origin',redirect:'error',cache:'no-store',...options,headers:{...headers,...options.headers}});
    const headers=await privateHeaders(),response=await send(headers);
    if(response.status!==403)return response;
    const previous=headers['X-Chaos-CSRF'];
    if(csrf===previous)csrf='';
    const refreshed=await privateHeaders();
    return refreshed['X-Chaos-CSRF']===previous?response:send(refreshed);
  }
  function resetWorkspace(){
    resetAISettings();
    resetAiSheet();
    resetLiveSync();
    clearSharedNavigation();
    checkedCards.clear();cardNetworks.clear();cardsWithSecurityCode.clear();clearPrivateDirectory();cardScope=null;closeDirectoryRemoval();cardEditorOpen=false;
    accountGroups.length=0;removedAccountKeys.clear();checkedAccounts.clear();expandedAccountEmails.clear();accountScope={retailer:null,group:null};closeAccountEditor();
    clearHierarchyUndo();
    for(const collection of [addresses,people,masterAccounts,masterCards,cardGroups,events])collection.length=0;
    ignoredAddressKeys.clear();cardIds.clear();cardLastFour.clear();cardBilling.clear();accountIds.clear();lastSources.clear();accountSerial=0;localAccountSerial=0;nextCardId=1;sourceSerial=0;exportLedger.length=0;
    clearTimeout(persistTimer);revision=0;lastAck='';conflicted=false;csrf='';pendingCards.clear();pendingPasswords.clear();hideCheckout();
    activeEvent=null;eventSerial=1;
    selectedPerson=null;selectedAddress=null;sourceMode='new-person';editedAccount=0;editedCard='';sourceTab='people';
    for(const node of root.querySelectorAll('input:not([type=checkbox]),textarea'))if(node.type!=='number'&&node.type!=='file')node.value='';
    closeMappedImport(true);resetTransfer();
    for(const node of [get('ec-status'),get('ec-source-status'),inv('account-msg'),inv('card-msg')])node.textContent='';
    updateRetailerOptions();drawCardGroups();drawAddressPicker();updatePalette();renderView();drawAccounts();drawCards();drawDirectory();
  }
  function updateRetailerOptions(){retailOptions();}
  let suiteSortDescending=false;
  function suiteRecipientsSummary(names){
    return names.length?'Recipients · '+names.map(name=>parseRecipientName(name).fullName).join(', '):'Recipients · use address names';
  }
  function suiteEditor(record){
    const saved=new Map((record.shippingVariants||[]).map(suite=>[suite.id,suite]));
    const suiteIds=[...saved.keys()];
    const locations=saved;
    const section=document.createElement('section');section.className='ec-shipping-destinations';
    section.innerHTML='<div class="ec-destinations-heading"><h3>Shipping destinations</h3></div><div class="ec-destination-chunks"></div>';
    const tools=document.createElement('div');tools.className='ec-destinations-tools';section.firstElementChild.append(tools);
    const items=section.querySelector('.ec-destination-chunks');
    const retainedRows=new Map();
    const sortSuites=()=>suiteIds.sort((a,b)=>compareShippingDestinations(locations.get(a),locations.get(b),suiteSortDescending)||a.localeCompare(b));
    sortSuites();
    if(suiteIds.length>1){
      const label=document.createElement('label');label.className='ec-suite-sort';label.textContent='Suite order';
      const order=document.createElement('select');order.setAttribute('aria-label','Suite order');
      order.append(new Option('Smallest first','asc'),new Option('Largest first','desc'));
      order.value=suiteSortDescending?'desc':'asc';label.append(order);tools.append(label);
      order.addEventListener('change',()=>{
        suiteSortDescending=order.value==='desc';
        sortSuites();
        if(!items.classList.contains('ec-people-window')){
          const list=items.firstElementChild,rows=new Map([...list.children].map(row=>[row.dataset.destinationId,row]));
          for(const key of suiteIds)list.append(rows.get(key));
          return;
        }
        for(const row of items.querySelectorAll('[data-destination-id]'))if(keepDraft(row))retainedRows.set(row.dataset.destinationId,row);
        stopSuiteWindow();items.scrollTop=0;drawSuites();
      });
    }
    if(!record.parcelMissing){
      const add=document.createElement('button');add.type='button';add.textContent='+ Destination';
      add.addEventListener('click',()=>startNewDestination(record.id));tools.append(add);
    }
    function destinationRow(key){
      const main=key===record.id,location=main?shippingLocation(record):locations.get(key);
      const title=main?'Main address':location.street2||location.street;
      const street=main||location.street2?streetAndUnit(location):'';
      const region=[location.state,location.zip].filter(Boolean).join(' ');
      const locality=[location.city,region,location.country].filter(Boolean).join(', ');
      const item=document.createElement('li');item.dataset.destinationId=key;
      const address=document.createElement('div');address.className='ec-destination-address';
      const heading=document.createElement('strong');heading.textContent=title;address.append(heading);
      for(const line of [street,locality])if(line){
        const span=document.createElement('span');span.textContent=line;address.append(span);
      }
      const label=document.createElement('label');label.className='ec-destination-switch';
      const toggle=document.createElement('input');toggle.type='checkbox';toggle.setAttribute('role','switch');toggle.dataset.destinationEnabled='';
      toggle.setAttribute('aria-label','Use '+title+' for shipping');
      toggle.checked=main?record.mainShippingEnabled!==false:saved.get(key)?.enabled===true;
      label.append(toggle,'Use for shipping');item.append(address,label);
      if(!main){
        const names=saved.get(key)?.recipients||[];
        const details=document.createElement('details');details.className='ec-destination-recipients';
        const summary=document.createElement('summary');summary.textContent=suiteRecipientsSummary(names);
        details.append(summary);
        details.insertAdjacentHTML('beforeend',recipientEntry('suite-recipient',names,'ec-suite-name-'+encodeURIComponent(key),'Recipient name for '+title));
        details.querySelector('[data-suite-recipient-add]').setAttribute('aria-label','Add recipient for '+title);
        item.append(details);
      }
      return item;
    }
    function keepDraft(block){
      return !!block.querySelector('details[open]')||[...block.querySelectorAll('[data-suite-recipient-input]')].some(input=>input.value);
    }
    function drawSuites(){
      stopSuiteWindow=mountPeopleChunks(items,[record.id,...suiteIds],{tag:'ul',size:6,estimate:130,keep:keepDraft,pinned:new Set(retainedRows.keys()),paint:(block,keys)=>{
        for(const key of keys){
          const row=retainedRows.get(key)||destinationRow(key);
          retainedRows.delete(key);block.append(row);
        }
      }});
    }
    drawSuites();
    section.addEventListener('change',event=>{
      const input=event.target;if(!input.matches('[data-destination-enabled]'))return;
      const key=input.closest('li').dataset.destinationId;
      if(key===record.id)record.mainShippingEnabled=input.checked;
      else saved.get(key).enabled=input.checked;
      reconcileShippingParent(record);drawAddressPicker();renderConfig(false);saveCoordinatorChange({immediate:true});
    });
    function changeName(item,removal){
      const input=item.querySelector('[data-suite-recipient-input]'),name=removal?removal.dataset.removeSuiteRecipient:input.value.trim();
      if(!name)return;
      const key=item.dataset.destinationId,existing=saved.get(key)?.recipients||[];
      if(removal){
        saved.get(key).recipients=existing.filter(value=>value!==name);
        if(input.dataset.editingRecipient===name)clearRecipientInput(input);
        removal.closest('.ec-chip').remove();
      }else{
        const next=commitRecipientInput(input,existing,'suite-recipient');
        if(!next||next===existing)return;
        saved.get(key).recipients=next;
      }
      item.querySelector('summary').textContent=suiteRecipientsSummary(saved.get(key).recipients);
      input.focus();refreshSourceRows(record.id);renderConfig(false);saveCoordinatorChange({immediate:true});
    }
    section.addEventListener('click',event=>{
      if(event.target.closest('[data-suite-recipient-add]'))changeName(event.target.closest('li'),null);
    });
    section.addEventListener('ec:confirmed-remove',event=>{
      if(event.detail.kind==='suite-recipient')changeName(event.target.closest('li'),event.target);
    });
    section.addEventListener('keydown',event=>{
      if(event.key==='Enter'&&!event.isComposing&&event.target.matches('[data-suite-recipient-input]')){
        event.preventDefault();changeName(event.target.closest('li'),null);
      }
    });
    return section;
  }
  function showShippingDestinations(){
    const record=addresses.find(address=>address.id===selectedAddress);
    const host=get('ec-source-editor');
    if(record&&sourceMode==='address'&&!host.querySelector('.ec-shipping-destinations'))host.append(suiteEditor(record));
  }
  function protectControls(on){
    for(const id of ['ec-coordinator-view','ec-sources','ec-app-settings'])get(id).inert=on;
    for(const tab of get('ec-source-tabs').querySelectorAll('#ec-show-coordinator,#ec-show-settings,[data-sources-tab]'))tab.disabled=on;
  }
  async function loadWorkspace(){
    if(loading)return;
    loading=true;reload.disabled=true;protectControls(true);
    notice.textContent='Loading workspace…';notice.hidden=false;
    const first=!active;
    try{
      profileScope=scope;
      if(first){
        await readState();
        const address=directoryAddresses()[0];
        selectedAddress=address?.id??null;
        selectedPerson=address?.personId??people[0]?.id??null;
        sourceMode=address?'address':selectedPerson?'person':'new-person';
        active=true;
        lastAck=JSON.stringify(snapshotState());workView(requestedView());startLiveSync();
      }else{
        await readState({ifClean:true});
        redrawSyncedWorkspace();
        if(!conflicted)await flushPersistence();
      }
      notice.textContent='';notice.hidden=true;
    }catch(error){
      notice.textContent=error.message||'Workspace unavailable. Use Refresh workspace to retry.';
    }finally{loading=false;reload.disabled=false;protectControls(!active);scheduleLiveUpdate();}
  }
  root.addEventListener('ec:directory-rendered',showShippingDestinations);
  // The browser may normalize an older workspace or defer an incoming update
// while a form is open. Compare each side to its own base before combining it.
function mergeWorkspace(localBase, local, remoteBase, remote, choice=null) {
  const conflicts=new Set(), collisions=[], absent=undefined;
  const object=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
  const own=(value,key)=>object(value)&&Object.hasOwn(value,key)?value[key]:absent;
  function equal(a,b){
    if(a===b)return true;
    if(Array.isArray(a))return Array.isArray(b)&&a.length===b.length&&a.every((value,i)=>equal(value,b[i]));
    if(!object(a)||!object(b))return false;
    const keys=Object.keys(a);
    return keys.length===Object.keys(b).length&&keys.every(key=>Object.hasOwn(b,key)&&equal(a[key],b[key]));
  }
  function upgrade(value){
    if(!value||value.version!==1)return value||{};
    const events=(value.events||[]).map(({sniped,shikari,...event})=>{
      const lists=[['sniped',sniped],['shikari',shikari]].map(([id,list])=>({
        basis:'person',...list,id,format:id,name:id==='sniped'?'Sniped':'Shikari'}));
      return {...event,lists,selectedListId:value.selected?.event===event.id?value.selected.bot:'sniped'};
    });
    return {...value,version:2,events,selected:{event:value.selected?.event??null}};
  }
  const lb=upgrade(localBase), ours=upgrade(local), rb=upgrade(remoteBase), saved=upgrade(remote);
  if(!remote&&!remoteBase)return {state:structuredClone(ours),conflicts:[]};
  const preferred=choice==='mine'?ours:saved;
  const identities={people:'Person',addresses:'Address',accounts:'Retail account',accountGroups:'Retail account group',
    cards:'Card',groups:'Card group',events:'Event',lists:'Profile list',ledger:'Export entry',shippingVariants:'Shipping destination'};
  const pairs=new Set(['accountKeys','sourceBaselines','names','edits']);
  const sets=new Set(['groupIds','disabledColumns','removedAccountKeys']);
  const allocation=['quantity','chosen','cards','cardGroupNames','useUngroupedCards','additionalCardLabels','accountGroupIds','pool','mode','basis','seed','retailer','rows','lastAssignments','nicknameCounters','names'];
  function clash(label,l,r){conflicts.add(label);return choice==='mine'?l:r;}
  function recipientKey(value){
    if(typeof value!=='string')return absent;
    const parts=value.split('|').map(part=>part.trim());
    const full=parts.length===2?parts.filter(Boolean).join(' '):value;
    return full.normalize('NFKC').trim().toLowerCase().replace(/\s+/gu,' ');
  }
  function mergeRecipient(a,l,b,r){
    const [al,lv,ar,rv]=[a,l,b,r].map(value=>typeof value==='string'?value.replace(/\s*\|\s*/g,'|'):value);
    const selected=merge(al,lv,ar,rv,'','Recipient name',true);
    return selected===rv?r:l;
  }
  function recipientOrder([al,lv,ar,rv]){
    // Editing the boundary retains the recipient's rotation slot. Honor an
    // independent reorder and keep an explicitly restored recipient in place.
    const sharedOrder=(first,second)=>[...first.keys()].filter(key=>second.has(key));
    const reordered=(base,current)=>!equal(sharedOrder(base,current),sharedOrder(current,base));
    const mine=reordered(al,lv),theirs=reordered(ar,rv);
    if(mine&&theirs&&!equal(sharedOrder(lv,rv),sharedOrder(rv,lv)))conflicts.add('Recipient order');
    const localFirst=mine===theirs?choice==='mine':mine;
    return localFirst?[lv,rv]:[rv,lv];
  }
  function collection(a,l,b,r,field,label,getKey,mergeItem){
    const inputs=[a,l,b,r].map(value=>value??[]);
    if(inputs.some(rows=>!Array.isArray(rows)))return clash(label,l,r);
    const maps=inputs.map(rows=>new Map(rows.map(item=>[getKey(item),item])));
    if(maps.some((map,i)=>map.size!==inputs[i].length||map.has(undefined)))return clash(label,l,r);
    const out=[];
    const [first,second]=field==='recipients'?recipientOrder(maps):[maps[3],maps[1]];
    for(const key of new Set([...first.keys(),...second.keys()])){
      const [al,lv,ar,rv]=maps.map(map=>map.get(key));
      let item;
      if(identities[field]&&al===absent&&ar===absent&&lv!==absent&&rv!==absent){
        const sameSource=lv.parcelKey&&lv.parcelKey===rv.parcelKey&&equal(lv,rv);
        if(sameSource)item=rv;
        else {
          collisions.push({field,key,l:lv,r:rv});
          item=clash('New '+label.toLowerCase()+' identity',lv,rv);
        }
      }else item=mergeItem(al,lv,ar,rv,key);
      if(item!==absent)out.push(item);
    }
    return out;
  }
  function merge(a,l,b,r,field='',label='Workspace',item=false){
    if(equal(a,l))return r;
    if(item&&(l===absent||r===absent)){
      if(l===r)return l;
      if(!equal(b,r)||l!==absent&&a!==absent)return clash(label+' removal',l,r);
      return l;
    }
    if(equal(a,b)&&equal(b,r))return l;
    if(field==='serials'&&object(l)&&object(r)){
      const keys=new Set([...Object.keys(r),...Object.keys(l)]);
      return Object.fromEntries([...keys].map(key=>[key,(key==='accountSerial'?Math.min:Math.max)(l[key]??r[key],r[key]??l[key])]));
    }
    if(Array.isArray(l)&&Array.isArray(r)){
      if(field==='recipients')return collection(a,l,b,r,field,'Recipient name',recipientKey,mergeRecipient);
      if(sets.has(field)||field==='cards'&&label==='Card group'){
        const [al,lv,ar,rv]=[a,l,b,r].map(values=>new Set(values||[]));
        return [...new Set([...r,...l])].filter(key=>merge(al.has(key),lv.has(key),ar.has(key),rv.has(key),'',label+' membership'));
      }
      if(field==='cards'&&label==='Profile list')return !equal(b,r)&&!equal(l,r)?clash(label+' cards',l,r):l;
      if(Object.hasOwn(identities,field))return collection(a,l,b,r,field,identities[field],v=>field==='groups'?v?.name:v?.id,
        (al,lv,ar,rv)=>merge(al,lv,ar,rv,'',identities[field],true));
      if(pairs.has(field))return collection(a,l,b,r,field,label,v=>Array.isArray(v)?v[0]:absent,(al,lv,ar,rv,key)=>{
        const value=merge(al?.[1],lv?.[1],ar?.[1],rv?.[1],'',field==='edits'?'Profile cell':label,true);
        return value===absent?absent:[key,value];
      });
    }
    if(object(l)&&object(r)){
      const keys=new Set([...Object.keys(r),...Object.keys(l),...Object.keys(a||{})]);
      const result=Object.create(null);
      const bundle=label==='Profile list'&&allocation.some(k=>!equal(own(a,k),own(l,k)))&&
        allocation.some(k=>!equal(own(b,k),own(r,k)))&&allocation.some(k=>!equal(own(l,k),own(r,k)));
      if(bundle)conflicts.add('Profile list assignments');
      for(const key of keys){
        const next=bundle&&allocation.includes(key)?own(choice==='mine'?l:r,key):
          merge(own(a,key),own(l,key),own(b,key),own(r,key),key,label);
        if(next!==absent)result[key]=next;
      }
      return result;
    }
    if(equal(l,r)||equal(b,r))return l;
    if(field==='selectedListId'||field==='event'&&label==='Workspace')return r;
    return clash(field==='name'?label+' name':label, l,r);
  }
  const state=structuredClone(merge(lb,ours,rb,saved));
  function selectRecord(field,key){
    const id=field==='groups'?'name':'id',from=preferred[field]?.find(row=>row[id]===key);
    const rows=state[field]||[],index=rows.findIndex(row=>row[id]===key);
    if(index>=0){if(from)rows[index]=structuredClone(from);else rows.splice(index,1);}
    else if(from)rows.push(structuredClone(from));
    state[field]=rows;
  }
  function selectPair(field,key){
    const values=new Map(state[field]||[]),source=new Map(preferred[field]||[]);
    if(source.has(key))values.set(key,source.get(key));else values.delete(key);
    state[field]=structuredClone([...values]);
  }
  // Card groups are keyed by their editable names. Treat a rename as one
  // action when a second client edits, deletes, or renames the old group.
  const groupMaps=[lb,ours,rb,saved].map(s=>new Map((s.groups||[]).map(g=>[g.name,g])));
  const [alGroups,localGroups,arGroups,remoteGroups]=groupMaps;
  function possibleRenames(base,current,old){
    if(current.has(old.name))return [];
    const added=[...current.values()].filter(group=>!base.has(group.name));
    const matches=added.filter(group=>equal(group.cards,old.cards));
    return (matches.length?matches:added).map(group=>group.name);
  }
  for(const [name,old] of alGroups){
    if(!arGroups.has(name))continue;
    const mine=possibleRenames(alGroups,localGroups,old);
    const theirs=possibleRenames(arGroups,remoteGroups,arGroups.get(name));
    if(!(mine.length&&!equal(arGroups.get(name),remoteGroups.get(name))||
         theirs.length&&!equal(old,localGroups.get(name))))continue;
    conflicts.add('Card group rename');
    for(const key of new Set([name,...mine,...theirs]))selectRecord('groups',key);
  }
  // If two unsynchronized clients allocated the same ID, references from
  // the losing record must never be attached to the winning identity.
  for(const {field,key,l,r} of collisions){
    const cards=[l.label,r.label],branches=[ours,saved];
    if(field==='people')for(const branch of branches)for(const row of branch.addresses||[]){
      if(row.personId===key)selectRecord('addresses',row.id);
    }
    if(field==='cards')for(const branch of branches)for(const row of branch.groups||[]){
      if(row.cards?.some(card=>cards.includes(card)))selectRecord('groups',row.name);
    }
    if(field==='accounts')for(const branch of branches)for(const [source,id] of branch.accountKeys||[]){
      if(id===key){selectPair('accountKeys',source);selectPair('sourceBaselines',source);}
    }
    if(field==='accountGroups')for(const branch of branches)for(const row of branch.accounts||[]){
      if(row.groupIds?.includes(key))selectRecord('accounts',row.id);
    }
    for(const branch of branches)for(const entry of branch.ledger||[]){
      const touches=field==='events'&&entry.eventId===key||
        (entry.profiles||[]).some(profile=>field==='accounts'&&profile.accountId===key||field==='cards'&&profile.cardId===key||
           field==='addresses'&&profile.addressId===key||field==='shippingVariants'&&profile.destinationId===key);
      if(touches)selectRecord('ledger',entry.id);
    }
    if(['cards','accounts','addresses','accountGroups','groups'].includes(field)){
      for(const branch of branches)for(const event of branch.events||[]){
        const touches=(event.lists||[]).some(list=>{
          if(field==='cards')return (list.cards||[]).some(card=>cards.includes(card));
          if(field==='groups')return list.cardGroupNames?.includes(key);
          if(field==='accountGroups')return list.accountGroupIds?.includes(key);
          if(field==='addresses')return (list.chosen||[]).some(id=>String(id).includes(key));
          return (list.rows||[]).some(row=>row[0]===key)||(list.edits||[]).some(pair=>pair[0]===key);
        });
        if(touches)selectRecord('events',event.id);
      }
    }
  }
  // Detect dependencies which are valid in both inputs but lose their target
  // through a concurrent deletion, rename, or assignment in the other tab.
  function destinations(addresses){
    const result=new Map();
    for(const address of addresses||[]){
      result.set(address.id,address.id);result.set('shipping-group:'+address.id,address.id);
      for(const suite of address.shippingVariants||[])result.set(suite.id,address.id);
    }
    return result;
  }
  const knownDestinations=new Map([...destinations(ours.addresses),...destinations(saved.addresses)]);
  function problems(s){
    const issues=new Map();
    const report=(label,owners)=>issues.set(JSON.stringify([label,...owners]),{label,owners});
    const names=new Map(),cardLabels=new Map(),groups=new Set((s.groups||[]).map(g=>g.name));
    const people=new Set((s.people||[]).map(p=>p.id)),accounts=new Map((s.accounts||[]).map(a=>[a.id,a]));
    const accountGroups=new Map((s.accountGroups||[]).map(g=>[g.id,g]));
    const currentDestinations=destinations(s.addresses);
    const norm=v=>String(v||'').trim().toLowerCase();
    function unique(token,field,id,label){
      if(names.has(token))report(label,[[field,names.get(token)],[field,id]]);
      else names.set(token,id);
    }
    for(const card of s.cards||[]){cardLabels.set(card.label,card.id);unique('card:'+norm(card.label),'cards',card.id,'Duplicate card label');}
    for(const group of s.accountGroups||[])unique(JSON.stringify(['group',group.retailer,norm(group.name)]),'accountGroups',group.id,'Duplicate retail account group');
    for(const acc of s.accounts||[]){
      for(const login of new Set([norm(acc.email),norm(acc.username)].filter(Boolean)))unique(JSON.stringify(['login',acc.retailer,login]),'accounts',acc.id,'Duplicate retail login');
      for(const id of acc.groupIds||[])if(accountGroups.get(id)?.retailer!==acc.retailer)report('Retail account group membership',[['accounts',acc.id],['accountGroups',id]]);
    }
    for(const address of s.addresses||[])if(address.personId&&!people.has(address.personId))report('Address owner',[['addresses',address.id],['people',address.personId]]);
    function checkCard(label,field,key){
      if(!label||cardLabels.has(label))return;
      const linked=(ours.cards||[]).find(c=>c.label===label)||(saved.cards||[]).find(c=>c.label===label);
      if(linked)report('Assigned card',[[field,key],['cards',linked.id]]);
    }
    for(const g of s.groups||[])for(const label of g.cards||[])checkCard(label,'groups',g.name);
    for(const event of s.events||[]){
      const used=new Map(),owner=['events',event.id];
      function checkDestination(id){
        if(knownDestinations.has(id)&&!currentDestinations.has(id))report('Assigned shipping destination',[owner,['addresses',knownDestinations.get(id)]]);
      }
      for(const list of event.lists||[]){
        for(const id of list.chosen||[])checkDestination(id);
        for(const label of list.cards||[])checkCard(label,'events',event.id);
        for(const g of list.cardGroupNames||[])if(!groups.has(g))report('Assigned card group',[owner,['groups',g]]);
        for(const g of list.accountGroupIds||[])if(!accountGroups.has(g))report('Assigned retail account group',[owner,['accountGroups',g]]);
        for(const [account,address,card] of list.rows||[]){
          if(!accounts.has(account)||list.retailer&&accounts.get(account)?.retailer!==list.retailer)report('Assigned retail account',[owner,['accounts',account]]);
          checkDestination(address);
          if(!event.sharing&&used.has(account)&&used.get(account)!==list.id)report('Retail account assigned to two lists',[owner]);
          used.set(account,list.id);checkCard(card,'events',event.id);
        }
      }
    }
    return issues;
  }
  const before=[problems(ours),problems(saved)],replaced=new Set();
  let changed=true;
  while(changed){
    changed=false;
    for(const [key,{label,owners}] of problems(state))if(before.every(issues=>!issues.has(key))){
      conflicts.add(label);
      for(const [field,id] of owners){
        const token=JSON.stringify([field,id]);
        if(replaced.has(token))continue;
        selectRecord(field,id);replaced.add(token);changed=true;
      }
    }
  }
  for(const event of state.events||[]){
    if(!event.lists?.some(list=>list.id===event.selectedListId))event.selectedListId=event.lists?.[0]?.id??null;
  }
  if(state.events&&(!state.selected||!state.events.some(event=>event.id===state.selected.event))){
    state.selected={event:state.events[0]?.id??null};
  }
  return {state,conflicts:[...conflicts]};
}

  const saveNote=document.createElement('span');
saveNote.className='ec-source-note';saveNote.setAttribute('role','status');reload.before(saveNote);
let inFlightDirty=false,checkoutToken=0,currentSecret=null,editingDetails=false;
function safeCopy(record){const {password,...clean}=record;return clean;}
function collectPasswordChanges(){
  for(const account of masterAccounts)if(account.password){pendingPasswords.set(String(account.id),account.password);account.hasPassword=true;account.password='';}
}
function packList(list){
  const groups=selectedCardGroups(list);
  return {id:list.id,name:list.name,format:list.format,quantity:list.quantity,chosen:[...list.chosen],destinationSelection:1,cards:[...list.cards],pool:list.pool,mode:list.mode,basis:list.basis,seed:list.seed,
    retailer:list.retailer||'Target',cardGroupNames:groups.map(group=>group.name),useUngroupedCards:list.useUngroupedCards,
    ...(groups.length||list.useUngroupedCards?{additionalCardLabels:additionalCardSelection(list,groups)}:{}),
    sort:list.sort,names:[...list.names],edits:[...list.edits],nicknameCounters:list.nicknameCounters,
    accountEmailColumn:true,disabledColumns:[...list.disabledColumns],accountGroupIds:list.accountGroupIds??null,lastAssignments:list.lastAssignments||[],rows:list.rows.map(row=>[row.accountId,row.addressId,row.card])};
}
function snapshotState(){
  return {version:2,people,addresses,accounts:masterAccounts.map(safeCopy),accountGroups,removedAccountKeys:[...removedAccountKeys],
    cards:masterCards.map(label=>({label,id:cardIds.get(label),lastFour:cardLastFour.get(label)||'',network:cardNetworks.get(cardIds.get(label))||'',billing:cardBilling.get(cardIds.get(label))||null})),
    groups:cardGroups,events:events.map(event=>({id:event.id,name:event.name,sharing:event.sharing,selectedListId:sharedSelectedList(event),lists:event.lists.map(packList)})),
    selected:{event:sharedSelectedEvent()},serials:{sourceSerial,eventSerial,localAccountSerial,accountSerial,nextCardId},
    accountKeys:[...accountIds],sourceBaselines:[...lastSources].map(([key,row])=>[key,safeCopy(row)]),ledger:exportLedger};
}
function savedList(saved){
  const list=makeList(saved.id,saved.name,saved.format);
  list.accountGroupIds=Array.isArray(saved.accountGroupIds)?saved.accountGroupIds:null;
  for(const key of ['quantity','pool','mode','basis','seed','retailer','sort'])if(key in saved)list[key]=saved[key];
  list.chosen=new Set(saved.destinationSelection?saved.chosen:saved.chosen.map(shippingGroupKey));
  list.cards=new Set(saved.cards);list.names=new Map(saved.names);list.edits=new Map(saved.edits);
  if(Array.isArray(saved.lastAssignments)){
    list.lastAssignments=saved.lastAssignments.filter(pair=>Array.isArray(pair)&&Number.isInteger(pair[0])&&typeof pair[1]==='string');
  }else{
    // Older lists released their assignments when the pool was turned off. Their
    // saved nickname entries still identify the account and destination for each edit.
    const prior=new Map();
    for(const [key] of list.names){
      if(typeof key!=='string')continue;
      const split=key.indexOf(':'),id=Number(key.slice(0,split));
      if(split>0&&Number.isInteger(id))prior.set(id,key.slice(split+1));
    }
    list.lastAssignments=[...prior].sort(([a],[b])=>Number(list.edits.has(b))-Number(list.edits.has(a)));
  }
  if(Array.isArray(saved.cardGroupNames))list.cardGroupNames=saved.cardGroupNames;
  list.useUngroupedCards=saved.useUngroupedCards===true;
  list.disabledColumns=new Set(Array.isArray(saved.disabledColumns)?saved.disabledColumns.filter(key=>fieldLabels.has(key)):[]);
  // A legacy list could hide the shipping contact while showing the assigned retail login.
  if(!saved.accountEmailColumn&&!list.disabledColumns.has('login'))list.disabledColumns.delete('email');
  list.nicknameCounters=Object.assign(Object.create(null),saved.nicknameCounters);
  const byAccount=new Map(masterAccounts.map(row=>[row.id,row])),byAddress=new Map(allShippingDestinations().map(row=>[row.id,row]));
  const counts=new Map();
  list.rows=saved.rows.flatMap(([id,addressId,card])=>{
    const account=byAccount.get(id),address=byAddress.get(addressId);
    return account&&address?[buildRow(list,account,address,card,nextAddressNumber(counts,addressId))]:[];
  });
  return list;
}
function unpackState(state){
  if(![1,2].includes(state.version)||!Array.isArray(state.events))throw Error('Unsupported saved workspace.');
  rememberSharedNavigation(state);
  clearHierarchyUndo();
  accountGroups.length=0;for(const group of state.accountGroups||[])accountGroups.push(group);
  removedAccountKeys.clear();for(const key of state.removedAccountKeys||[])removedAccountKeys.add(key);
  for(const [target,rows] of [[people,state.people],[addresses,state.addresses],[masterAccounts,state.accounts],[cardGroups,state.groups],[exportLedger,state.ledger||[]]]){target.length=0;for(const row of rows)target.push(row);}
  masterCards.length=0;cardIds.clear();cardLastFour.clear();cardBilling.clear();cardNetworks.clear();cardsWithSecurityCode.clear();invalidateDirectorySecrets();for(const card of state.cards){masterCards.push(card.label);cardIds.set(card.label,card.id);cardLastFour.set(card.label,card.lastFour);if(card.hasSecurityCode)cardsWithSecurityCode.add(card.id);if(card.network)cardNetworks.set(card.id,card.network);if(card.billing)cardBilling.set(card.id,card.billing);}
  accountIds.clear();for(const [key,id] of state.accountKeys)accountIds.set(key,id);
  lastSources.clear();for(const [key,row] of state.sourceBaselines)lastSources.set(key,row);
  // Earlier releases copied every Parcel routing account into the retail inventory.
  // Preserve deliberate retail edits and keep untouched copies available to saved profile assignments.
  const parcelGroupIds=new Set(accountGroups.filter(g=>normalized(g.name)==='parcel mail'&&normalized(g.retailer)==='target'&&
    masterAccounts.some(a=>a.parcelKey&&!a.formerParcelGroups&&a.groupIds?.includes(g.id))).map(g=>g.id));
  for(let i=accountGroups.length-1;i>=0;i--)if(parcelGroupIds.has(accountGroups[i].id))accountGroups.splice(i,1);
  for(const account of masterAccounts){
    const former=account.groupIds?.filter(id=>parcelGroupIds.has(id));
    if(former?.length){
      account.formerParcelGroups=[...new Set([...(account.formerParcelGroups||[]),...former])];
      account.groupIds=account.groupIds.filter(id=>!parcelGroupIds.has(id));
      // This membership alone qualified mirrors as managed under the old eligibility rule.
      if(account.parcelKey)delete account.retailManaged;
    }
    if(!account.parcelKey||account.retailManaged)continue;
    const baseline=lastSources.get(account.parcelKey);
    if(account.hasPassword||Object.hasOwn(secretVersions.passwords||{},String(account.id))||account.username||account.groupIds?.length||
      baseline&&['retailer','email'].some(field=>(baseline[field]||'')!==(account[field]||'')))account.retailManaged=true;
  }
  ({sourceSerial,eventSerial,localAccountSerial,accountSerial,nextCardId}=state.serials);
  events.length=0;
  const changedPools=[],loose=ungroupedCardLabels();
  for(const saved of state.events){
    const event=makeEvent(saved.id,saved.name);
    event.sharing=!!saved.sharing;
    const legacy=state.version===1;
    const lists=legacy?['sniped','shikari'].map(format=>({basis:'person',...saved[format],id:format,name:formatLabel(format),format})):saved.lists;
    event.lists=lists.map(config=>{
      const list=savedList(config);
      // Concurrent saves can merge a newly selected pool with newer directory membership.
      if(Array.isArray(config.additionalCardLabels)){
        const selected=cardsFromPools(selectedCardGroups(list),list.useUngroupedCards?loose:[],config.additionalCardLabels);
        if(setCardPoolMembers(list,selected))changedPools.push([event,list]);
      }
      return list;
    });
    const selected=legacy?(event.id===state.selected.event?state.selected.bot:'sniped'):saved.selectedListId;
    event.selectedListId=event.lists.some(list=>list.id===selected)?selected:event.lists[0]?.id||null;
    events.push(event);
    eventSerial=Math.max(eventSerial,Number(event.id.slice(6))+1);
  }
  for(const [event,list] of changedPools)rebuildList(event,list);
  activeEvent=events.some(event=>event.id===state.selected.event)?state.selected.event:events[0]?.id||null;
  restoreTabNavigation();
  drawCardGroups();renderEvents();
}
function workspaceHasPendingChanges(){
  return !!saving||pendingCards.size>0||pendingPasswords.size>0||masterAccounts.some(account=>account.password)||lastAck!==JSON.stringify(snapshotState());
}
async function fetchWorkspaceState(){
  const response=await fetch('/api/event-coordinator/state',{credentials:'same-origin',cache:'no-store',redirect:'error',mode:'same-origin',headers:{'X-Parcel-Workspace':scope||profileScope}});
  if(!response.ok)throw requestError(response.status,'Could not open workspace.');
  const body=await response.json();
  if(body.profileId!==profileScope||!Number.isSafeInteger(body.revision))throw requestError(409,'Workspace changed.');
  return body;
}
async function readState({ifClean=false}={}){
  if(ifClean&&(saving||!canApplyRemoteView()||workspaceHasPendingChanges()))return;
  const body=await fetchWorkspaceState();
  if(ifClean&&(saving||!canApplyRemoteView()||workspaceHasPendingChanges()||body.revision<=revision))return;
  acceptWorkspace(body);
  if(ifClean)lastAck=JSON.stringify(snapshotState());
  saveNote.textContent=revision?'Saved':'';
}
function saveFailureMessage(status){
  if(status===401||status===403)return 'Save paused. Refresh the workspace and retry.';
  if(status===413)return 'The workspace exceeds the save size limit. Reduce the batch and retry.';
  if(status===400||status===422)return 'Save rejected. Check the edited fields and retry.';
  return 'Save interrupted. Your edits are still here. Choose Refresh workspace to retry.';
}
function refreshPersistence(){
  if(!active||conflicted)return;inFlightDirty=true;clearTimeout(persistTimer);
  persistTimer=setTimeout(()=>{flushPersistence().catch(()=>{});},300);
}
saveCoordinatorChange=({immediate=false}={})=>{
  if(!active||conflicted)return;
  refreshPersistence();
  if(immediate)flushPersistence().catch(()=>{});
};
async function savePendingWorkspace(){
  for(let attempt=0;attempt<4;attempt++){
    collectPasswordChanges();
    const packed=JSON.stringify(snapshotState()),cards=new Map(pendingCards),passwords=new Map(pendingPasswords);
    inFlightDirty=false;
    if(lastAck===packed&&!cards.size&&!passwords.size)return;
    saveNote.textContent='Saving…';
    const response=await privateRequest('/api/event-coordinator/state',{method:'PUT',
      body:'{"revision":'+revision+',"state":'+packed+',"changes":'+JSON.stringify({cards:Object.fromEntries(cards),passwords:Object.fromEntries(passwords)})+'}'});
    if(response.status===409){
      const reason=await response.json().catch(()=>({}));
      if(reason.detail==='The workspace changed. Reload Event Manager.')throw requestError(403,'Reload Event Manager.');
      const latest=await fetchWorkspaceState();
      if(latest.revision<=revision)throw Error('Could not confirm the saved revision.');
      if(!reconcileWorkspace(latest))throw requestError(409,'Waiting to sync.');
      continue;
    }
    if(!response.ok)throw requestError(response.status,saveFailureMessage(response.status));
    const result=await response.json();
    if(result.profileId!==profileScope||result.revision!==revision+1)throw Error('Could not confirm the saved revision.');
    revision=result.revision;lastAck=packed;acknowledgedState=JSON.parse(packed);secretVersions=result.secretRevisions||{};
    for(const [id,value] of cards)if(pendingCards.get(id)===value)pendingCards.delete(id);
    for(const [id,value] of passwords)if(pendingPasswords.get(id)===value)pendingPasswords.delete(id);
    saveNote.textContent='Saved';return;
  }
  throw Error('Another editor is saving. Retry shortly.');
}
async function flushPersistence(){
  clearTimeout(persistTimer);if(!active)return;
  if(conflicted)throw requestError(409,'Resolve the edits shown in this tab.');
  if(saving){await saving;if(inFlightDirty)return flushPersistence();return;}
  saving=savePendingWorkspace();
  try{await saving;}catch(error){
    if(error.status!==409)saveNote.textContent=saveFailureMessage(error.status);
    inFlightDirty=true;throw error;
  }finally{saving=null;}
  if(inFlightDirty)return flushPersistence();
  scheduleLiveUpdate();
}
for(const name of ['input','change','submit','click'])root.addEventListener(name,event=>{
  if(loading||!active||event.target===reload||event.target.closest('.ec-private-transient'))return;
  refreshPersistence();
});
window.addEventListener('beforeunload',event=>{
  if(!active)return;collectPasswordChanges();
  if(saving||pendingCards.size||pendingPasswords.size||lastAck!==JSON.stringify(snapshotState())){event.preventDefault();event.returnValue='';}
});
document.addEventListener('visibilitychange',()=>{if(document.hidden&&active&&inFlightDirty)flushPersistence().catch(()=>{});});
async function fetchSecrets(cards=[],accounts=[],directory=false){
  for(let attempt=0;attempt<2;attempt++){
    if(!directory)await flushPersistence();
    const requestedRevision=revision;
    const response=await privateRequest('/api/event-coordinator/secrets',{method:'POST',
      body:JSON.stringify({revision:requestedRevision,cards,accounts:accounts.map(String)})});
    if(response.ok)return response.json();
    if(response.status===401||response.status===403)throw requestError(response.status,'Refresh the workspace to open saved details.');
    if(response.status!==409)break;
    const reason=await response.json().catch(()=>({}));
    if(reason.detail==='The workspace changed. Reload Event Manager.')throw requestError(403,'Reload Event Manager.');
    if(saving)await saving.catch(()=>{});
    if(requestedRevision===revision){
      signalNewerRevision(requestedRevision+1);
      throw requestError(409,'Workspace changed. Refresh workspace to open the current saved details.');
    }
  }
  throw Error('Could not open current saved details.');
}
reloadDirectorySources=()=>{checkServerRevision();scheduleLiveUpdate();};
directoryHasPendingChanges=()=>workspaceHasPendingChanges()||!canApplyRemoteView();
readDirectorySecrets=async(kind,records)=>{
  const ids=records.map(record=>String(record.id));
  // A failed write elsewhere in the sheet must not block a read of saved credentials.
  // Keep edits staged during the request visible even if their save completes first.
  if(kind==='cards'){
    const staged=new Map(ids.map(id=>[id,pendingCards.get(id)]));
    const saved=(await fetchSecrets(ids,[],true)).cards;
    return new Map(ids.map(id=>[id,{...saved[id],...staged.get(id),...pendingCards.get(id)}]));
  }
  const staged=new Map(records.map(record=>[String(record.id),record.password||pendingPasswords.get(String(record.id))]));
  const saved=(await fetchSecrets([],ids,true)).passwords;
  return Object.fromEntries(records.map(record=>{
    const id=String(record.id);
    return [id,record.password||(pendingPasswords.get(id)??staged.get(id)??saved[id]??'')];
  }));
};
resolveCheckout=async ids=>new Map(Object.entries((await fetchSecrets(ids)).cards||{}));
stageImportedCard=(id,details)=>{
  noteCardChange(id,details);pendingCards.set(id,{...pendingCards.get(id),...details});
};
forgetCardSecrets=id=>pendingCards.delete(id);
finishInventoryImport=async kind=>{
  if(kind==='cards')hideCheckout();
  await flushPersistence();
};
readRetailPasswords=async records=>records.length?readDirectorySecrets('accounts',records):{};
forgetAccountSecrets=records=>{
  for(const account of records){
    pendingPasswords.delete(String(account.id));
    if(account.parcelKey)lastSources.delete(account.parcelKey);
  }
};
const detailsPanel=document.createElement('div'),detailsToggle=document.createElement('button'),detailsFields=document.createElement('div');
detailsToggle.type='button';detailsToggle.className='cursor-interaction';detailsToggle.textContent='Edit checkout details';detailsFields.hidden=true;
detailsPanel.className='ec-private-transient';detailsFields.className='ec-source-fields';detailsPanel.append(detailsToggle,detailsFields);
const checkFields=[['holder','Cardholder'],['month','Expiry month (MM)'],['year','Expiry year (YYYY)'],['brand','Brand'],['paymentType','Payment method'],['firstName','Billing first name'],['lastName','Billing last name'],['street','Billing street'],['street2','Billing street 2'],['city','Billing city'],['state','Billing state'],['zip','Billing ZIP'],['country','Billing country']];
const checkoutInputs=new Map();
for(const [key,label] of checkFields){
  const row=document.createElement('label'),input=document.createElement('input');input.type='text';input.autocomplete='off';
  if(['month','year'].includes(key))input.inputMode='numeric';input.maxLength=500;row.append(label,input);detailsFields.append(row);checkoutInputs.set(key,input);
}
inv('card-form').insertBefore(detailsPanel,inv('card-form').lastElementChild);
function hideCheckout(){
  checkoutToken++;editingDetails=false;currentSecret=null;detailsFields.hidden=true;detailsToggle.textContent='Edit checkout details';detailsToggle.disabled=false;
  for(const input of checkoutInputs.values()){input.value='';input.setCustomValidity('');input.placeholder='';cardDraftInputs.delete(input);}
}
detailsToggle.addEventListener('click',async()=>{
  if(editingDetails){
    for(const [id,key] of [['card-number','number'],['card-code','cvv']]){
      const input=inv(id);delete input.dataset.edited;input.value=currentSecret?.[key]||'';input.setCustomValidity('');
    }
    showCardNumberIssue('');hideCheckout();return;
  }
  const token=++checkoutToken,id=cardIds.get(editedCard);detailsToggle.disabled=true;
  try{
    const saved=id?(await readDirectorySecrets('cards',[{id}])).get(id):null;if(token!==checkoutToken)return;
    currentSecret={...(saved||{})};editingDetails=true;detailsFields.hidden=false;detailsToggle.textContent='Discard checkout edits';
    for(const [key,input] of checkoutInputs)input.value=saved?.[key]||'';
    fillPrivateField(inv('card-number'),saved?.number);fillPrivateField(inv('card-code'),saved?.cvv);
    updateNetworkHint();
  }catch{if(token===checkoutToken)inv('card-msg').textContent='Could not open checkout details.';}finally{if(token===checkoutToken)detailsToggle.disabled=false;}
});
function updateNetworkHint(){
  const number=inv('card-number').value.replace(/[\s-]/g,'')||currentSecret?.number;
  const detected=detectedNetwork(number);
  checkoutInputs.get('brand').placeholder=detected?'Detected · '+detected:'Card network';
}
detailsFields.addEventListener('input',event=>event.target.setCustomValidity(''));
inv('card-number').addEventListener('input',updateNetworkHint);
const checkedCardSubmissions=new WeakMap();
inv('card-form').addEventListener('submit',event=>{
  if(!editingDetails)return;
  const result={...currentSecret};
  delete result.cvv;
  const number=inv('card-number').value.replace(/[\s-]/g,'');
  if(number)result.number=number;
  for(const [key,input] of checkoutInputs)result[key]=input.value.trim();
  const invalid=!/^\d{12,19}$/.test(result.number||'')?'number':!/^(0?[1-9]|1[0-2])$/.test(result.month)?'month':!/^\d{4}$/.test(result.year)?'year':'';
  if(invalid){event.preventDefault();event.stopImmediatePropagation();const input=invalid==='number'?inv('card-number'):checkoutInputs.get(invalid);input.setCustomValidity(invalid==='number'?'Enter 12–19 card digits.':invalid==='month'?'Enter a month from 01 to 12.':'Enter the four-digit year.');input.reportValidity();return;}
  const label=inv('card-name').value.trim(),original=editedCard;
  if(!label||masterCards.some(card=>card!==original&&normalized(card)===normalized(label)))return;
  result.month=result.month.padStart(2,'0');inv('four').value=result.number.slice(-4);
  checkedCardSubmissions.set(event,{label,result});
},true);
inv('card-form').addEventListener('submit',event=>{
  const checked=checkedCardSubmissions.get(event);if(!checked)return;
  const {label,result}=checked,immutableId=cardIds.get(label);if(!immutableId||editedCard!==label)return;
  stageImportedCard(immutableId,result);
  const network=displayNetwork(result.brand,result.number);
  if(network)cardNetworks.set(immutableId,network);else cardNetworks.delete(immutableId);
  const billing=Object.fromEntries(['firstName','lastName','street','street2','city','state','zip','country'].map(key=>[key,result[key]||'']));
  if(Object.values(billing).some(Boolean))cardBilling.set(immutableId,billing);else cardBilling.delete(immutableId);
  refreshSourceRows();drawCardDirectory();renderConfig(false);hideCheckout();refreshPersistence();
});
root.addEventListener('ec:card-editor',hideCheckout);
get('ec-show-coordinator').addEventListener('click',hideCheckout);root.addEventListener('ec:show-source',hideCheckout);

  // Keep the server baseline separate from the tab's normalized sheet snapshot.
let acknowledgedState=null,secretVersions={},newestRevision=0,conflictRevision=null;
let refreshJob=null,liveTimer=0,checkTimer=0,stream=null,syncGeneration=0,liveStarted=false,streamConfirmed=false;
const conflictBar=document.createElement('div');
conflictBar.id='ec-sync-conflict';conflictBar.className='ec-source-actions';conflictBar.hidden=true;
conflictBar.innerHTML='<span role="status"></span><button type="button" id="ec-sync-keep">Keep my edits</button><button type="button" id="ec-sync-latest">Use saved edits</button>';
notice.after(conflictBar);
const conflictMessage=conflictBar.querySelector('span');

function clearConflict(){
  conflicted=false;conflictRevision=null;conflictBar.hidden=true;conflictMessage.textContent='';
}
function resetLiveSync(){
  syncGeneration++;liveStarted=false;stream?.close();stream=null;streamConfirmed=false;
  clearTimeout(liveTimer);clearTimeout(checkTimer);newestRevision=0;refreshJob=null;
  acknowledgedState=null;secretVersions={};clearConflict();
}
function installSharedState(state){
  unpackState(structuredClone(state));
}
function acceptWorkspace(body){
  secretVersions=body.secretRevisions||{};
  if(body.state)installSharedState(body.state);
  acknowledgedState=body.state;revision=body.revision;
  newestRevision=Math.max(newestRevision,revision);clearConflict();
}
function protectPrivateEdits(base,local,body,plan,choice){
  const discard=[],labels=new Set();
  let canKeep=true;
  const loginFields=['retailer','email','username','parcelKey'];
  const differentLogin=(left,right)=>loginFields.some(key=>(left?.[key]??'')!==(right?.[key]??''));
  for(const [records,kind,pending] of [['accounts','passwords',pendingPasswords],['cards','cards',pendingCards]]){
    const old=new Map(base[records].map(row=>[String(row.id),row]));
    const serverBase=new Map((acknowledgedState?.[records]||base[records]).map(row=>[String(row.id),row]));
    const mine=new Map(local[records].map(row=>[String(row.id),row]));
    const saved=new Map((body.state?.[records]||[]).map(row=>[String(row.id),row]));
    const before=secretVersions[kind]||{},after=body.secretRevisions?.[kind]||{};
    const noun=kind==='cards'?'card':'retail account';
    for(const id of new Set([...old.keys(),...mine.keys(),...pending.keys()])){
      const createdTwice=!old.has(id)&&mine.has(id)&&saved.has(id);
      const removedThere=old.has(id)&&!saved.has(id)&&mine.has(id);
      const removedHere=old.has(id)&&!mine.has(id)&&saved.has(id);
      const secretChanged=(before[id]||0)!==(after[id]||0);
      const privateConflict=secretChanged&&(pending.has(id)||removedHere);
      const editedDeletedRecord=removedThere&&(pending.has(id)||(before[id]&&JSON.stringify(old.get(id))!==JSON.stringify(mine.get(id))));
      const loginConflict=kind==='passwords'&&old.has(id)&&mine.has(id)&&saved.has(id)&&differentLogin(mine.get(id),saved.get(id))&&(
        pending.has(id)&&differentLogin(serverBase.get(id),saved.get(id))||secretChanged&&differentLogin(old.get(id),mine.get(id)));
      if(!createdTwice&&!editedDeletedRecord&&!privateConflict&&!loginConflict)continue;
      if(createdTwice||removedThere){
        // The old ID now identifies a different record, or its private data was
        // deleted. Never attach a partial secret patch to that identity.
        canKeep=false;labels.add(createdTwice?'New '+noun+' identity':'Deleted '+noun);
      }else if(loginConflict){
        labels.add('Account login and password');
        if(!pending.has(id))canKeep=false;
        if(choice){
          const record=plan.state.accounts.find(row=>String(row.id)===id);
          const chosen=(choice==='saved'?saved:mine).get(id);
          if(record)for(const field of loginFields){
            if(field in chosen)record[field]=chosen[field];else delete record[field];
          }
        }
      }else labels.add(kind==='cards'?'Private card details':'Account password');
      if(choice==='saved'){
        discard.push([pending,id]);
        if(createdTwice||removedHere||removedThere){
          plan.state[records]=plan.state[records].filter(row=>String(row.id)!==id);
          if(saved.has(id))plan.state[records].push(saved.get(id));
        }
      }
    }
  }
  return {discard,labels:[...labels],canKeep};
}
function reconcileWorkspace(body,choice=null){
  if(body.revision<=revision)return true;
  newestRevision=Math.max(newestRevision,body.revision);
  collectPasswordChanges();
  const encoded=JSON.stringify(snapshotState()),dirty=encoded!==lastAck||pendingCards.size||pendingPasswords.size;
  if(!dirty){
    if(!canApplyRemoteView()){saveNote.textContent='Changes available · finish editing to sync';return false;}
    acceptWorkspace(body);lastAck=JSON.stringify(snapshotState());redrawSyncedWorkspace();
    saveNote.textContent='Up to date';return true;
  }
  const local=JSON.parse(encoded),base=JSON.parse(lastAck);
  const plan=mergeWorkspace(base,local,acknowledgedState||base,body.state||base,choice);
  const secrets=protectPrivateEdits(base,local,body,plan,choice);
  const collisions=[...new Set([...plan.conflicts,...secrets.labels])];
  if(collisions.length&&(!choice||choice==='mine'&&!secrets.canKeep)){
    conflicted=true;conflictRevision=body.revision;
    conflictMessage.textContent='Overlapping changes · '+collisions.slice(0,3).join(', ')+
      (collisions.length>3?' +'+(collisions.length-3):'')+
      (secrets.canKeep?'. Choose which edits to use.':'. Use the saved record and re-enter your changes for the correct card or account.');
    get('ec-sync-keep').hidden=!secrets.canKeep;conflictBar.hidden=false;
    saveNote.textContent='Review edits';return false;
  }
  if(!choice&&!canApplyRemoteView()){
    saveNote.textContent='Changes available · finish editing to sync';return false;
  }
  for(const [pending,id] of secrets.discard)pending.delete(id);
  acceptWorkspace(body);
  lastAck=JSON.stringify(snapshotState());
  installSharedState(plan.state);
  for(const [id,patch] of pendingCards)noteCardChange(id,patch);
  for(const account of masterAccounts)if(pendingPasswords.has(String(account.id)))account.hasPassword=!!pendingPasswords.get(String(account.id));
  redrawSyncedWorkspace();saveNote.textContent='Up to date';return true;
}
async function resolveSyncConflict(choice){
  if(!conflicted||saving||refreshJob)return;
  const approved=conflictRevision;
  const buttons=conflictBar.querySelectorAll('button');for(const button of buttons)button.disabled=true;
  try{
    const latest=await fetchWorkspaceState();
    // A decision about one pair of edits does not authorize overwriting a later edit.
    const updatedAgain=latest.revision!==approved;
    if(!reconcileWorkspace(latest,updatedAgain?null:choice)){
      if(updatedAgain&&conflicted)conflictMessage.textContent='Another edit just arrived. '+conflictMessage.textContent;
      return;
    }
    await flushPersistence();
  }catch(error){if(error.status!==409)saveNote.textContent=saveFailureMessage(error.status);}
  finally{for(const button of buttons)button.disabled=false;scheduleLiveUpdate();}
}
get('ec-sync-keep').addEventListener('click',()=>resolveSyncConflict('mine'));
get('ec-sync-latest').addEventListener('click',()=>resolveSyncConflict('saved'));
function signalNewerRevision(value){
  if(!Number.isSafeInteger(value)||value<=revision)return;
  newestRevision=Math.max(newestRevision,value);scheduleLiveUpdate();
}
function scheduleLiveUpdate(){
  if(!active||!liveStarted||document.hidden||newestRevision<=revision||conflicted)return;
  clearTimeout(liveTimer);liveTimer=setTimeout(pullLiveUpdate,70);
}
async function pullLiveUpdate(){
  if(!active||!liveStarted||document.hidden||loading||saving||refreshJob||conflicted||newestRevision<=revision)return;
  if(!canApplyRemoteView()){
    saveNote.textContent='Changes available · finish editing to sync';return;
  }
  const generation=syncGeneration;
  const job=fetchWorkspaceState();refreshJob=job;
  try{
    const body=await job;
    if(generation!==syncGeneration||loading||saving||conflicted||body.revision<=revision)return;
    if(reconcileWorkspace(body)&&(lastAck!==JSON.stringify(snapshotState())||pendingCards.size||pendingPasswords.size))refreshPersistence();
  }catch{if(generation===syncGeneration&&!saving&&!conflicted)saveNote.textContent='Sync interrupted · retrying';}
  finally{if(refreshJob===job)refreshJob=null;}
}
async function checkServerRevision(){
  if(!active||!liveStarted||document.hidden)return;
  const generation=syncGeneration;
  try{
    const response=await fetch('/api/event-coordinator/revision',{credentials:'same-origin',mode:'same-origin',redirect:'error',cache:'no-store',headers:{'X-Parcel-Workspace':profileScope}});
    if(response.ok){
      const body=await response.json();
      if(generation===syncGeneration&&body.profileId===profileScope)signalNewerRevision(body.revision);
    }
  }catch{/* The next connection or foreground check retries. */}
  if(generation===syncGeneration&&liveStarted&&!document.hidden)armRevisionCheck();
}
function armRevisionCheck(){
  clearTimeout(checkTimer);
  checkTimer=setTimeout(checkServerRevision,streamConfirmed&&stream?.readyState===1?30000:2000);
}
function startLiveSync(){
  liveStarted=true;if(!active||document.hidden)return;
  if(!stream&&typeof EventSource==='function'){
    const connection=new EventSource('/api/event-coordinator/events?workspace='+encodeURIComponent(profileScope));stream=connection;streamConfirmed=false;
    connection.onmessage=event=>{
      if(connection!==stream)return;
      try{const body=JSON.parse(event.data);if(body.profileId===profileScope&&Number.isSafeInteger(body.revision)){streamConfirmed=true;signalNewerRevision(body.revision);}}catch{}
      armRevisionCheck();
    };
    connection.onerror=()=>{if(stream===connection){streamConfirmed=false;armRevisionCheck();}};
    connection.addEventListener('login',()=>{
      if(stream!==connection)return;
      connection.close();stream=null;saveNote.textContent='Reload the workspace to resume sync';armRevisionCheck();
    });
  }
  armRevisionCheck();scheduleLiveUpdate();
}
function pauseLiveSync(){
  clearTimeout(checkTimer);clearTimeout(liveTimer);stream?.close();stream=null;streamConfirmed=false;
}
function resumeLiveSync(){
  if(!liveStarted||document.hidden)return;
  startLiveSync();checkServerRevision();
  if(!loading&&!saving&&!conflicted&&workspaceHasPendingChanges())refreshPersistence();
}
document.addEventListener('visibilitychange',()=>{if(document.hidden)pauseLiveSync();else resumeLiveSync();});
window.addEventListener('pagehide',pauseLiveSync);
for(const name of ['pageshow','focus','online'])window.addEventListener(name,resumeLiveSync);
for(const name of ['focusout','click','change','submit','keyup','close'])root.addEventListener(name,scheduleLiveUpdate,true);

  const keyPanel=get('ec-app-settings'),keyInput=get('ec-ai-key'),keySave=get('ec-ai-key-save');
const keyRemove=get('ec-ai-key-remove'),keyStatus=get('ec-ai-key-status');
let keyConfigured=false,keyBusy=false,keyRequest=0,keyRemovalTimer=0;

function cancelKeyRemoval(){
  clearTimeout(keyRemovalTimer);keyRemovalTimer=0;keyRemove.textContent='Remove key';
}
function updateKeyControls(){
  get('ec-ai-key-badge').hidden=!keyConfigured;
  get('ec-ai-key-label').textContent=keyConfigured?'Replacement API key':'API key';
  keySave.textContent=keyConfigured?'Replace key':'Save key';
  keySave.disabled=keyBusy||!keyInput.value.trim();
  keyRemove.disabled=keyBusy;keyRemove.hidden=!keyConfigured;
  keyInput.disabled=keyBusy;
}
function resetAISettings(){
  keyRequest++;keyBusy=false;keyConfigured=false;keyInput.value='';keyStatus.textContent='';
  cancelKeyRemoval();updateKeyControls();
}
function keyRequestMessage(status){
  if(status===401||status===403||status===409)return 'Reload the workspace to manage the API key.';
  if(status===400||status===422)return 'Check the API key and try again.';
  return 'Could not reach the key settings. Try again.';
}
async function readKeyStatus(){
  if(keyPanel.hidden||!active||keyBusy)return;
  const token=++keyRequest;
  try{
    const response=await fetch('/api/event-coordinator/ai/settings',{credentials:'same-origin',mode:'same-origin',redirect:'error',cache:'no-store',headers:{'X-Parcel-Workspace':scope||profileScope}});
    if(!response.ok)throw requestError(response.status);
    const result=await response.json();
    if(token!==keyRequest)return;
    keyConfigured=result.configured===true;keyStatus.textContent='';
    updateKeyControls();
  }catch(error){if(token===keyRequest)keyStatus.textContent=keyRequestMessage(error.status);}
}
async function changeKey(method){
  if(!active||keyBusy||keyPanel.hidden)return;
  const token=++keyRequest;
  const options={method};
  if(method==='PUT'){
    const apiKey=keyInput.value.trim();
    if(!apiKey)return;
    options.body=JSON.stringify({apiKey});
  }
  cancelKeyRemoval();keyBusy=true;updateKeyControls();
  keyStatus.textContent=method==='PUT'?'Saving key…':'Removing key…';
  try{
    const response=await privateRequest('/api/event-coordinator/ai/settings',options);
    if(!response.ok)throw requestError(response.status);
    const result=await response.json();
    if(token!==keyRequest)return;
    keyConfigured=result.configured===true;keyInput.value='';
    keyStatus.textContent=keyConfigured?'API key saved.':'API key removed.';
  }catch(error){if(token===keyRequest)keyStatus.textContent=keyRequestMessage(error.status);}
  finally{if(token===keyRequest){keyBusy=false;updateKeyControls();}}
}
get('ec-show-settings').hidden=false;
get('ec-ai-key-form').addEventListener('submit',event=>{event.preventDefault();changeKey('PUT');});
keyInput.addEventListener('input',()=>{cancelKeyRemoval();updateKeyControls();});
keyRemove.addEventListener('click',()=>{
  if(keyRemovalTimer){changeKey('DELETE');return;}
  keyRemove.textContent='Confirm removal';keyRemovalTimer=setTimeout(cancelKeyRemoval,2000);
});
root.addEventListener('ec:workspace-view',event=>{
  cancelKeyRemoval();
  if(event.detail==='settings'){readKeyStatus();return;}
  keyInput.value='';updateKeyControls();
});
window.addEventListener('focus',readKeyStatus);
window.addEventListener('pagehide',()=>{keyInput.value='';cancelKeyRemoval();updateKeyControls();});

  // OpenAI describes edits using column IDs. Interpret and preview them against local rows.
const aiDialog=get('ec-ai-dialog'),aiScope=get('ec-ai-scope'),aiInstruction=get('ec-ai-instruction');
const aiPreview=get('ec-ai-preview'),aiApply=get('ec-ai-apply'),aiStatus=get('ec-ai-status');
let aiList=null,aiPromptList='',aiSelection=new Map(),aiPlan=null,aiRequest=null;
const aiChangesTable=get('ec-ai-changes');
let aiRowsObserver=null;
function clearAiPreview(){
  aiPlan=null;aiApply.disabled=true;get('ec-ai-results').hidden=true;
  aiRowsObserver?.disconnect();aiRowsObserver=null;
  for(const body of Array.from(aiChangesTable.tBodies))body.remove();
  get('ec-ai-settings').hidden=true;aiStatus.textContent='';
}
function updateAiPrompt(){
  aiRequest?.abort();aiRequest=null;clearAiPreview();
  aiPreview.disabled=!aiInstruction.value.trim()||!sheetRowsByKey.size||!aiColumns().length;
}
function resetAiSheet(){
  aiRequest?.abort();aiRequest=null;aiList=null;aiPromptList='';aiSelection.clear();
  aiInstruction.value='';clearAiPreview();
  if(aiDialog.open)aiDialog.close();
}
function aiColumns(){
  const selected=new Set([...aiSelection.values()].flatMap(fields=>[...fields]));
  return visibleColumns.map(([field])=>field).filter(field=>editableFields.has(field)&&!flagFields.has(field)
    &&(aiScope.value!=='cells'||selected.has(field)));
}
get('ec-ai-edit').hidden=false;
get('ec-ai-edit').addEventListener('click',()=>{
  aiList=currentList();if(!active||!aiList)return;
  const key=JSON.stringify([profileScope,activeEvent,aiList.id]);
  if(aiPromptList!==key)aiInstruction.value='';
  aiPromptList=key;aiSelection.clear();
  let count=0;
  for(const [row,fields] of selectedCells){
    if(!sheetRowsByKey.has(row))continue;
    const textFields=new Set([...fields].filter(field=>editableFields.has(field)&&!flagFields.has(field)));
    if(textFields.size){aiSelection.set(row,textFields);count+=textFields.size;}
  }
  aiScope.replaceChildren();
  if(selectedCells.size)aiScope.add(new Option('Selected cells ('+count+' editable text '+(count===1?'cell':'cells')+')','cells'));
  aiScope.add(new Option('All profiles in this list ('+sheetRowsByKey.size+')','list'));
  updateAiPrompt();
  if(!sheetRowsByKey.size)aiStatus.textContent='Add profiles to this list to start editing.';
  else if(!aiColumns().length)aiStatus.textContent='Select an editable text cell or choose all profiles.';
  aiDialog.showModal();aiInstruction.focus();
});
aiScope.addEventListener('change',updateAiPrompt);
aiInstruction.addEventListener('input',updateAiPrompt);
get('ec-ai-close').addEventListener('click',()=>aiDialog.close());
aiDialog.addEventListener('close',()=>{
  aiRequest?.abort();aiRequest=null;aiSelection.clear();aiList=null;clearAiPreview();
});
get('ec-ai-settings').addEventListener('click',()=>{aiDialog.close();get('ec-show-settings').click();});

const aiOperations=new Set(['set','prefix','suffix','replace','trim','uppercase','lowercase','titlecase','extract_email','copy','prefix_letters']);
// Match the coordinator's 12 MiB save limit. Check expanding replacements before allocating them.
const aiSizeLimit=12*1024*1024;
function aiSizeError(){return Error('This instruction would exceed the workspace size limit. Use a shorter replacement or select fewer cells.');}
function validateAiSteps(steps,fields){
  if(!Array.isArray(steps)||steps.length>12||steps.some(step=>!step||!fields.includes(step.field)||!aiOperations.has(step.op)
      ||typeof step.value!=='string'||step.value.length>500||typeof step.find!=='string'||step.find.length>500
      ||step.op==='copy'&&!fields.includes(step.value)||step.op==='replace'&&!step.find
      ||step.op==='prefix_letters'&&!/^[1-8]$/.test(step.value)))throw Error('The suggested edit could not be applied to these columns. Try a more specific instruction.');
}
function aiValue(step,row){
  const current=String(row[step.field]??'');
  switch(step.op){
    case 'set':return step.value;
    case 'prefix':return step.value+current;
    case 'suffix':return current+step.value;
    case 'replace':{
      const growth=step.value.length-step.find.length;
      if(growth>0){
        let size=current.length,start=0;
        while((start=current.indexOf(step.find,start))!==-1){
          size+=growth;
          if(size>aiSizeLimit)throw aiSizeError();
          start+=step.find.length;
        }
      }
      return current.replaceAll(step.find,()=>step.value);
    }
    case 'trim':return current.trim();
    case 'uppercase':return current.toUpperCase();
    case 'lowercase':return current.toLowerCase();
    case 'titlecase':return current.toLowerCase().replace(/\p{L}[\p{L}\p{M}]*/gu,word=>word[0].toUpperCase()+word.slice(1));
    case 'extract_email':{
      for(const domain of current.matchAll(/@[a-z0-9-]+(?:\.[a-z0-9-]+)+/gi)){
        const local=current.slice(Math.max(0,domain.index-64),domain.index).match(/[a-z0-9.!#$%&'*+/=?^_`{|}~-]+$/i);
        if(local)return local[0]+domain[0];
      }
      return current;
    }
    case 'copy':return String(row[step.value]??'');
    case 'prefix_letters':{
      const bytes=crypto.getRandomValues(new Uint8Array(Number(step.value)));
      return Array.from(bytes,value=>String.fromCharCode(65+value%26)).join('')+' '+current;
    }
  }
}
function buildAiPreview(steps){
  const simulated={edits:new Map(aiList.edits)},changes=[],overrides=new Map();
  let storedCharacters=0;
  const columns=visibleColumns.filter(([field])=>editableFields.has(field)&&!flagFields.has(field));
  for(const [id,row] of sheetRowsByKey){
    const selected=aiSelection.get(id);
    if(aiScope.value==='cells'&&!selected)continue;
    const key=row.draftKey??row.accountId,copy={...row};
    if(key==null)continue;
    for(const step of steps){
      if(aiScope.value==='cells'&&!selected.has(step.field))continue;
      const value=aiValue(step,copy);
      if(value.length>aiSizeLimit)throw aiSizeError();
      if(String(copy[step.field]??'')!==value)applySheetEdits(simulated,copy,{[step.field]:value});
    }
    const diffs=columns.filter(([field])=>String(row[field]??'')!==String(copy[field]??''));
    if(!diffs.length)continue;
    const edited=simulated.edits.get(key);
    for(const field of Object.keys(edited))if(typeof edited[field]==='string')storedCharacters+=edited[field].length;
    if(storedCharacters>aiSizeLimit)throw aiSizeError();
    overrides.set(key,edited);
    for(const [field,label] of diffs)changes.push([row.nickname,label,String(row[field]??''),String(copy[field]??'')]);
  }
  return {overrides,changes};
}
function drawAiPreview(changes){
  const panel=get('ec-ai-results'),blockSize=40;
  panel.hidden=!changes.length;panel.scrollTop=0;
  const windowed=changes.length>2*blockSize&&typeof IntersectionObserver==='function';
  aiChangesTable.toggleAttribute('data-windowed',windowed);
  aiChangesTable.setAttribute('aria-rowcount',String(changes.length+1));
  function showRows(body){
    const start=Number(body.dataset.start),rows=document.createDocumentFragment();
    for(let index=start;index<Math.min(start+blockSize,changes.length);index++){
      const tr=document.createElement('tr');
      tr.setAttribute('aria-rowindex',String(index+2));
      for(const value of changes[index]){const td=tr.insertCell();td.textContent=value||'—';}
      rows.append(tr);
    }
    body.replaceChildren(rows);
  }
  function reserveHeight(body,height){
    const space=document.createElement('tr');space.className='ec-ai-spacer';space.setAttribute('aria-hidden','true');
    const td=space.insertCell();td.colSpan=4;td.style.height=Math.ceil(height)+'px';
    body.replaceChildren(space);
  }
  if(windowed){
    aiRowsObserver=new IntersectionObserver((entries,observer)=>{
      if(observer!==aiRowsObserver)return;
      for(const entry of entries){
        const body=entry.target,spacer=body.firstElementChild.classList.contains('ec-ai-spacer');
        if(entry.isIntersecting&&spacer)showRows(body);
        else if(!entry.isIntersecting&&!spacer)reserveHeight(body,entry.boundingClientRect.height);
      }
    },{root:panel,rootMargin:'240px 0px'});
  }
  for(let start=0;start<changes.length;start+=blockSize){
    const body=aiChangesTable.createTBody();body.dataset.start=String(start);
    if(!windowed||start===0)showRows(body);
    else reserveHeight(body,Math.min(blockSize,changes.length-start)*32);
    aiRowsObserver?.observe(body);
  }
}
async function requireAiRevision(expected,request){
  const response=await fetch('/api/event-coordinator/revision',{credentials:'same-origin',mode:'same-origin',redirect:'error',cache:'no-store',
    signal:request.signal,headers:{'X-Parcel-Workspace':profileScope}});
  if(!response.ok)throw Error('Could not check the current sheet. Reconnect, then try again.');
  const info=await response.json();
  if(info.profileId!==profileScope||!Number.isSafeInteger(info.revision))throw Error('Reload Event Manager for the current workspace.');
  signalNewerRevision(info.revision);
  if(!active||conflicted||info.revision!==expected||revision!==expected||newestRevision>expected||currentList()!==aiList){
    throw Error('The sheet changed. Close this dialog to sync, then preview the edit again.');
  }
}
aiPreview.addEventListener('click',async()=>{
  const instruction=aiInstruction.value.trim(),fields=aiColumns();
  if(!instruction||!fields.length||!aiList||aiRequest)return;
  clearAiPreview();aiPreview.disabled=true;aiStatus.textContent='Preparing preview…';
  const request=new AbortController();aiRequest=request;
  try{
    await flushPersistence();
    if(request.signal.aborted)return;
    const expected=revision;
    await requireAiRevision(expected,request);
    const response=await privateRequest('/api/event-coordinator/ai/plan',{method:'POST',signal:request.signal,
      body:JSON.stringify({instruction,fields})});
    const plan=await response.json().catch(()=>({}));
    if(request.signal.aborted)return;
    if(!response.ok){
      get('ec-ai-settings').hidden=response.status!==409&&!(typeof plan.detail==='string'&&plan.detail.includes('Settings'));
      throw Error(typeof plan.detail==='string'?plan.detail:'Could not prepare the edit. Try again.');
    }
    await requireAiRevision(expected,request);
    if(request.signal.aborted)return;
    validateAiSteps(plan.operations,fields);
    const {changes,overrides}=buildAiPreview(plan.operations);
    drawAiPreview(changes);
    aiStatus.textContent=(plan.message?plan.message+' ':'')+(changes.length?changes.length+' changed '+(changes.length===1?'cell':'cells')+' across '+overrides.size+' '+(overrides.size===1?'profile.':'profiles.'):'No cells would change.');
    if(changes.length){
      aiPlan={revision:expected,overrides,state:JSON.stringify(snapshotState())};aiApply.disabled=false;
    }
  }catch(error){
    if(!request.signal.aborted)aiStatus.textContent=error instanceof TypeError?'Could not reach Event Coordinator. Reconnect, then try again.':error.message;
  }finally{
    if(aiRequest===request){aiRequest=null;aiPreview.disabled=false;}
  }
});
aiApply.addEventListener('click',async()=>{
  const plan=aiPlan;
  if(!plan||aiRequest)return;
  const request=new AbortController();aiRequest=request;aiApply.disabled=true;aiPreview.disabled=true;
  try{
    await requireAiRevision(plan.revision,request);
    if(request.signal.aborted)return;
    if(plan.state!==JSON.stringify(snapshotState()))throw Error('The sheet changed. Preview the edit again.');
    const list=aiList;
    trackSheetEdit(list,plan.overrides.keys(),()=>{
      for(const [key,overrides] of plan.overrides)setSheetOverrides(list,key,overrides);
    });
    aiDialog.close();redrawSheetEdits(list);
  }catch(error){
    if(!request.signal.aborted){clearAiPreview();aiStatus.textContent=error instanceof TypeError?'Could not check the sheet. Reconnect, then preview again.':error.message;}
  }finally{
    if(aiRequest===request){aiRequest=null;aiPreview.disabled=false;}
  }
});

const exportTarget = get('ec-export-target');
const canPickDirectory = typeof window.showDirectoryPicker === 'function';
const canPickFile = typeof window.showSaveFilePicker === 'function';
let selectedDirectory = null;
exportTarget.innerHTML = '<label>Export destination<select id="ec-destination">'+
  '<option value="device">'+(canPickDirectory?'Choose a folder on this device…':canPickFile?'Choose a file on this device…':'Download on this device')+'</option></select></label><small id="ec-export-location-note"></small>';
function exportLocationNote(){
  get('ec-export-location-note').textContent=
    canPickDirectory||canPickFile?'Your browser will ask where to save the file.':'Uses your browser’s download location. Enable “Ask where to save” in browser settings to choose a folder.';
}
get('ec-destination').addEventListener('change',exportLocationNote);exportLocationNote();
async function writeExportHandle(handle,content){
  const stream=await handle.createWritable();
  try{await stream.write(content);await stream.close();}catch(error){await stream.abort().catch(()=>{});throw error;}
  return handle.name;
}
async function writeExportInDirectory(folder,filename,content){
  const dot=filename.lastIndexOf('.'),stem=filename.slice(0,dot),suffix=filename.slice(dot);
  for(let copy=1;;copy++){
    const candidate=copy===1?filename:stem+' ('+copy+')'+suffix;
    try{await folder.getFileHandle(candidate);}
    catch(error){
      if(error.name==='TypeMismatchError')continue;
      if(error.name!=='NotFoundError')throw error;
      return writeExportHandle(await folder.getFileHandle(candidate,{create:true}),content);
    }
  }
}
privateExportFile=async function(filename,content,kind,mime){
  if(canPickDirectory){
    selectedDirectory=await window.showDirectoryPicker({id:'event-coordinator-export',mode:'readwrite',...(selectedDirectory?{startIn:selectedDirectory}:{})});
    const folder=selectedDirectory,write=()=>writeExportInDirectory(folder,filename,content);
    return navigator.locks?navigator.locks.request('event-coordinator-folder-export',write):write();
  }else if(canPickFile){
    const type={description:{xlsx:'Excel workbook',json:'AYCD native JSON',csv:'CSV'}[kind],accept:{[mime]:['.'+kind]}};
    const handle=await window.showSaveFilePicker({id:'event-coordinator-export',suggestedName:filename,types:[type]});
    if((await handle.getFile()).size)throw new DOMException('Choose a new file name to keep the existing file.','ExportFileExistsError');
    return writeExportHandle(handle,content);
  }else{
    const url=URL.createObjectURL(new Blob([content],{type:mime}));
    const link=document.createElement('a');
    link.href=url;link.download=filename;document.body.append(link);link.click();link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),30000);
    return filename;
  }
};

  resetWorkspace();reload.addEventListener('click',loadWorkspace);loadWorkspace();
})();
return;

  restore(window.openai?.widgetState);renderView();
  window.addEventListener('openai:set_globals',e=>{if(e.detail?.globals?.widgetState){restore(e.detail.globals.widgetState);renderView();}});
})();
