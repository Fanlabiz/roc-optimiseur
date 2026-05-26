// ==UserScript==
// @name         Rise of Cultures - Export & Optimiseur
// @namespace    http://tampermonkey.net/
// @version      6.4
// @description  Collecte les batiments, optimise le placement, genere un fichier Excel
// @match        https://*.riseofcultures.com/*
// @grant        none
// @run-at       document-start
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// ==/UserScript==

(function () {
    'use strict';

    // ═══════════════════════════════════════════════════════════════
    // SECTION 1 : DÉCODAGE PROTOBUF (repris de v5.4)
    // ═══════════════════════════════════════════════════════════════

    function readVarint(bytes, offset) {
        var result = 0, shift = 0, b;
        do {
            if (offset >= bytes.length) return { value: result >>> 0, offset: offset };
            b = bytes[offset++];
            result |= (b & 0x7F) << shift;
            shift += 7;
        } while (b & 0x80);
        return { value: result >>> 0, offset: offset };
    }

    function decodeMsg(bytes, offset, end) {
        offset = offset || 0;
        end = (end !== undefined) ? end : bytes.length;
        var fields = {};
        while (offset < end) {
            var tagInfo = readVarint(bytes, offset);
            if (tagInfo.value === 0 || tagInfo.offset >= end) break;
            offset = tagInfo.offset;
            var fieldNum = tagInfo.value >>> 3;
            var wireType = tagInfo.value & 0x07;
            if (!fields[fieldNum]) fields[fieldNum] = [];
            if (wireType === 0) {
                var vi = readVarint(bytes, offset); offset = vi.offset;
                fields[fieldNum].push({ type: 'v', val: vi.value });
            } else if (wireType === 1) {
                fields[fieldNum].push({ type: 'f64' }); offset += 8;
            } else if (wireType === 2) {
                var lenInfo = readVarint(bytes, offset); offset = lenInfo.offset;
                var len = lenInfo.value;
                if (len < 0 || offset + len > bytes.length) break;
                fields[fieldNum].push({ type: 'b', data: bytes.slice(offset, offset + len) });
                offset += len;
            } else if (wireType === 5) {
                fields[fieldNum].push({ type: 'f32', data: bytes.slice(offset, offset + 4) }); offset += 4;
            } else { break; }
        }
        return fields;
    }

    function str(bytes) { try { return new TextDecoder('utf-8').decode(bytes); } catch(e) { return ''; } }
    function isText(b) {
        if (!b||b.length===0||b.length>500) return false;
        for (var i=0;i<b.length;i++){var c=b[i];if(c<9||(c>13&&c<32))return false;}
        return true;
    }
    function objToBytes(obj) {
        if (obj instanceof Uint8Array) return obj;
        if (obj instanceof ArrayBuffer) return new Uint8Array(obj);
        var keys = Object.keys(obj).map(Number).sort(function(a,b){return a-b;});
        var arr = new Uint8Array(keys.length);
        keys.forEach(function(k,i){ arr[i] = obj[k] & 0xFF; });
        return arr;
    }
    function readFloat32(data) {
        if (!data || data.length < 4) return 0;
        var buf = new ArrayBuffer(4);
        var view = new DataView(buf);
        for (var i = 0; i < 4; i++) view.setUint8(i, data[i]);
        return view.getFloat32(0, true);
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 2 : DONNÉES GLOBALES
    // ═══════════════════════════════════════════════════════════════

    var CATALOG = {}, LOCA = {}, STARTUP_CULTURE = {}, EVOLVING_DATA = {};
    var startupLoaded = false, catalogLoaded = false, locaLoaded = false;
    var lastStartupBytes = null, pendingCityBytes = [];
    var allBuildings = [], cities = {}, allLockedBlocks = {}, lockedBlocksKeys = {};
    var fullLog = [], logs = [];
    var modalPrios = {}, modalMode = 'priority';
    var _rocTimer = null;

    var EVOLVING_ERA_NAMES = {
        'BronzeAge':'BronzeAge','MinoanEra':'MinoanEra','ClassicGreece':'ClassicGreece',
        'EarlyRome':'EarlyRome','RomanEmpire':'RomanEmpire','ByzantineEra':'ByzantineEra',
        'AgeOfTheFranks':'AgeOfTheFranks','FeudalAge':'FeudalAge','IberianEra':'IberianEra',
        'KingdomOfSicily':'KingdomOfSicily','HighMiddleAges':'HighMiddleAges',
        'EarlyGothicEra':'EarlyGothicEra','LateGothicEra':'LateGothicEra'
    };
    var DYNAMIC_THRESHOLDS = {
        'DynamicAge_Farm_Rural':    [1470,2940,5880], 'DynamicAge_Home_Small':    [1200,2420,4820],
        'DynamicAge_Home_Average':  [1360,2730,5460], 'DynamicAge_Home_Premium':  [1280,2570,5130],
        'DynamicAge_Farm_Domestic': [1960,3920,7740], 'DynamicAge_Farm_Premium':  [1960,3920,7740],
        'StoneAge_Home_Premium':    [1280,2570,5130]
    };
    var RANGE_FORMULAS = {
        'DynamicAge_CultureSite_Large':[[0,3]],'DynamicAge_CultureSite_Moderate':[[0,2]],
        'DynamicAge_CultureSite_Compact':[[0,1]],'DynamicAge_CultureSite_Little':[[0,1]]
    };
    // Valeurs exactes des sites culturels par niveau (depuis app.py)
    var CULTURE_SITE_DATA = {
        'CultureSite_Large':   {14:[3890,3],10:[2800,3],7:[2100,3],5:[1540,3],3:[1050,3],1:[560,3]},
        'CultureSite_Moderate':{14:[1440,2],10:[1050,2],7:[770,2], 5:[560,2], 3:[385,2], 1:[210,2]},
        'CultureSite_Compact': {14:[800,1], 10:[560,1], 7:[420,1], 5:[315,1], 3:[210,1], 1:[105,1]},
        'CultureSite_Little':  {14:[390,1], 10:[280,1], 7:[210,1], 5:[154,1], 3:[105,1], 1:[56,1]}
    };
    var FIXED_CULTURAL = {
        'QuintennialMonument':[500,3],'SchoolV2':[0,0],'ArchitectsStudioV2':[0,0],
        'CommanderTower':[0,0],'BirdHouse':[0,0],'Saddler':[0,0],'TikiStatue':[0,0],
        'BasketDryingRack':[0,0],'BanyanTree':[0,0],'TempleGate':[0,0],
        'MidnightClock':[0,0],'WinterPavillion':[0,0],'AirshipExhibit':[0,0]
    };

    // ═══════════════════════════════════════════════════════════════
    // SECTION 3 : PARSING CATALOGUE & LOCA (repris de v5.4)
    // ═══════════════════════════════════════════════════════════════

    function loadCatalog(callback) {
        var req = indexedDB.open('/idbfs');
        req.onsuccess = function(e) {
            var db = e.target.result;
            db.transaction('FILE_DATA','readonly').objectStore('FILE_DATA').getAllKeys().onsuccess = function(ev) {
                var gdKey = null, locaKey = null;
                ev.target.result.forEach(function(k){
                    var ks = String(k);
                    if (ks.indexOf('GameDesignResponse.data') !== -1) gdKey = k;
                    if (ks.indexOf('LocaResponse.data') !== -1) locaKey = k;
                });
                var pending = 2;
                function done() { pending--; if (pending === 0) { catalogLoaded = true; locaLoaded = true; callback(); } }
                if (gdKey) {
                    db.transaction('FILE_DATA','readonly').objectStore('FILE_DATA').get(gdKey).onsuccess = function(ev2) {
                        var gdBytes = objToBytes(ev2.target.result.contents);
                        parseCatalog(gdBytes); parseEvolvingBuildings(gdBytes);
                        log('Catalogue: '+Object.keys(CATALOG).length+' bâtiments','#888'); done();
                    };
                } else { done(); }
                if (locaKey) {
                    db.transaction('FILE_DATA','readonly').objectStore('FILE_DATA').get(locaKey).onsuccess = function(ev2) {
                        parseLoca(objToBytes(ev2.target.result.contents));
                        log('Traductions: '+Object.keys(LOCA).length+' entrées',Object.keys(LOCA).length>0?'#a8e6a3':'#f08080'); done();
                    };
                } else { log('LocaResponse.data absent','#f08080'); done(); }
            };
        };
        req.onerror = function(){ catalogLoaded=true; locaLoaded=true; callback(); };
    }

    function parseLoca(bytes) {
        try {
            var root = decodeMsg(bytes);
            function strRaw(data){try{return new TextDecoder('utf-8').decode(data);}catch(e){return '';}}
            function isLocaKey(s){return s&&(s.indexOf('Base.')===0||s.indexOf('Tutorial.')===0);}
            function tryParseEntries(items){
                var count=0;
                items.forEach(function(item){
                    if(item.type!=='b') return;
                    try{
                        var entry=decodeMsg(item.data),s1='',s2='';
                        (entry[1]||[]).forEach(function(f){if(f.type==='b'&&f.data.length>0&&f.data.length<2000)s1=strRaw(f.data);});
                        (entry[2]||[]).forEach(function(f){if(f.type==='b'&&f.data.length>0&&f.data.length<5000)s2=strRaw(f.data);});
                        if(s1&&s2){if(isLocaKey(s1)){LOCA[s1]=s2;count++;}else if(isLocaKey(s2)){LOCA[s2]=s1;count++;}}
                    }catch(e){}
                });
                return count;
            }
            var found=0;
            Object.keys(root).forEach(function(fn){
                var items=root[fn];
                if(items.length===1&&items[0].type==='b'&&items[0].data.length<=10) return;
                found+=tryParseEntries(items);
                if(found===0){items.forEach(function(item){if(item.type!=='b')return;try{var sub=decodeMsg(item.data);Object.keys(sub).forEach(function(sfn){found+=tryParseEntries(sub[sfn]||[]);});}catch(e){}});}
            });
        } catch(e){log('Erreur parseLoca: '+e,'#f08080');}
    }

    function getLocaName(fullName) {
        if (!fullName) return '';
        var exactKey = 'Base.Buildings.'+fullName+'_Name';
        if (LOCA[exactKey]) return LOCA[exactKey];
        var noNum = fullName.replace(/_\d+$/,'');
        if (LOCA['Base.Buildings.'+noNum+'_Name']) return LOCA['Base.Buildings.'+noNum+'_Name'];
        var prefix='Base.Buildings.'+noNum, keys=Object.keys(LOCA);
        for(var i=0;i<keys.length;i++){if(keys[i].indexOf(prefix)===0&&keys[i].indexOf('_Name')!==-1)return LOCA[keys[i]];}
        var parts=fullName.replace(/^Building_/,'').split('_');
        var eras=['StoneAge','BronzeAge','MinoanEra','ClassicGreece','EarlyRome','RomanEmpire',
                  'ByzantineEra','AgeOfTheFranks','FeudalAge','IberianEra','KingdomOfSicily',
                  'HighMiddleAges','EarlyGothicEra','LateGothicEra','DynamicAge',
                  'Egypt','Mayas','Vikings','Celts','Mongols','Mali','Persian','Thai','Polynesia',
                  'EventHalloween','EventWinter','EventAztec','EventCelts','EventGreek','EventMongols',
                  'EventWorldFair','EventPolynesia','EventPersian','Harbor'];
        var coreParts=parts.filter(function(p){return !eras.some(function(e){return p===e;})&&!/^\d+$/.test(p)&&p!=='Premium'&&p!=='Average'&&p!=='Rural'&&p!=='Domestic';});
        if(coreParts.length>0){var coreStr=coreParts.join('_');for(var j=0;j<keys.length;j++){if(keys[j].indexOf(coreStr)!==-1&&keys[j].indexOf('_Name')!==-1&&keys[j].indexOf('Base.Buildings.')===0)return LOCA[keys[j]];}}
        return '';
    }

    function parseCatalog(bytes) {
        try {
            var root=decodeMsg(bytes);
            (root[2]||[]).forEach(function(item){
                if(item.type!=='b') return;
                try{
                    var anyMsg=decodeMsg(item.data),typeUrl='';
                    (anyMsg[1]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data))typeUrl=str(f.data);});
                    if(typeUrl.indexOf('BuildingDefinition')===-1) return;
                    (anyMsg[2]||[]).forEach(function(vf){
                        if(vf.type!=='b') return;
                        var msg=decodeMsg(vf.data),bName='',w=0,h=0,culture=0,range=0;
                        (msg[1]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data)){var s2=str(f.data);if(s2.startsWith('Building_'))bName=s2;}});
                        (msg[3]||[]).forEach(function(f){if(f.type==='v')w=f.val;});
                        (msg[4]||[]).forEach(function(f){if(f.type==='v')h=f.val;});
                        (msg[5]||[]).forEach(function(compItem){
                            if(compItem.type!=='b') return;
                            try{var anyComp=decodeMsg(compItem.data),compType='';
                                (anyComp[1]||[]).forEach(function(f){if(f.type==='b')compType=str(f.data);});
                                if(compType.indexOf('CultureComponent')===-1) return;
                                (anyComp[2]||[]).forEach(function(vf2){if(vf2.type!=='b')return;var cultMsg=decodeMsg(vf2.data);
                                    (cultMsg[2]||[]).forEach(function(f){if(f.type==='v')range=f.val;});
                                    (cultMsg[3]||[]).forEach(function(f){if(f.type==='v')culture=f.val;});
                                });
                            }catch(e2){}
                        });
                        var thresholds={};
                        (msg[11]||[]).forEach(function(f11item){if(f11item.type!=='b')return;try{var t=decodeMsg(f11item.data),ts=0;(t[1]||[]).forEach(function(f){if(f.type==='v')ts=f.val;});if(ts>0)thresholds[ts]=ts;}catch(e){}});
                        var tKeys=Object.keys(thresholds).map(Number).sort(function(a,b){return a-b;});
                        if(w>0&&h>0)CATALOG[bName]={w:w,h:h,culture:culture,range:range,t25:tKeys[0]||0,t50:tKeys[1]||0,t100:tKeys[2]||0};
                    });
                }catch(e){}
            });
        }catch(e){}
    }

    function applyFormula(formula,level){
        if(!formula)return null;
        var m;
        m=formula.match(/\(#level \* ([\d.]+)\) \+ ([\d.]+)/);if(m)return parseFloat(m[1])*level+parseFloat(m[2]);
        m=formula.match(/\(#level \+ ([\d.]+)\) \* ([\d.]+)/);if(m)return(level+parseFloat(m[1]))*parseFloat(m[2]);
        m=formula.match(/\(#level \* ([\d.]+)\)/);if(m)return parseFloat(m[1])*level;
        m=formula.match(/\(#level \/ ([\d.]+)\)$/);if(m)return level/parseFloat(m[1]);
        return null;
    }

    function parseEvolvingBuildings(bytes){
        try{
            var root=decodeMsg(bytes);
            (root[2]||[]).forEach(function(item){
                if(item.type!=='b')return;
                try{
                    var anyOuter=decodeMsg(item.data),typeUrl='';
                    (anyOuter[1]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data))typeUrl=str(f.data);});
                    if(typeUrl.indexOf('DynamicFloatValueDefinitionDTO')===-1)return;
                    (anyOuter[2]||[]).forEach(function(payloadField){
                        if(payloadField.type!=='b')return;
                        try{
                            var defMsg=decodeMsg(payloadField.data),dvName='';
                            (defMsg[1]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data))dvName=str(f.data);});
                            var isCulture=dvName.indexOf('_CultureValues')!==-1,isRange=dvName.indexOf('_CultureRange')!==-1;
                            if(!isCulture&&!isRange)return;
                            if(isCulture&&dvName.indexOf('_CultureValues_StoneAge')!==-1)return;
                            // Extract building name
                            var bName=null;
                            var m2=dvName.match(/Evolving_([A-Za-z0-9]+)_\d+_Culture/);if(m2)bName=m2[1];
                            if(!bName){m2=dvName.match(/Collectable_([A-Za-z0-9]+)_\d+_Culture/);if(m2)bName=m2[1];}
                            if(!bName){m2=dvName.match(/Building_[A-Za-z0-9]+_([A-Za-z0-9]+)_\d+_Culture/);if(m2)bName=m2[1];}
                            if(!bName)return;
                            var eraKey='range';
                            if(isCulture){var eraSuffix=dvName.replace(/.*_CultureValues_?/,'');eraKey=(!eraSuffix||!EVOLVING_ERA_NAMES[eraSuffix])?'DEFAULT':eraSuffix;}
                            if(!EVOLVING_DATA[bName])EVOLVING_DATA[bName]={};
                            if(!EVOLVING_DATA[bName][eraKey])EVOLVING_DATA[bName][eraKey]={discrete:{},formula:null};
                            (defMsg[2]||[]).forEach(function(anyWrap){
                                if(anyWrap.type!=='b')return;
                                try{
                                    var awm=decodeMsg(anyWrap.data);
                                    (awm[2]||[]).forEach(function(cf){
                                        if(cf.type!=='b')return;
                                        try{
                                            var change=decodeMsg(cf.data);
                                            (change[1]||[]).forEach(function(entry){
                                                if(entry.type!=='b')return;
                                                try{var e=decodeMsg(entry.data),lvStr='';(e[1]||[]).forEach(function(f){if(f.type==='b')lvStr=str(f.data);});var lv=parseInt(lvStr);if(!lv)return;
                                                    (e[2]||[]).forEach(function(anyVal){if(anyVal.type!=='b')return;try{var av=decodeMsg(anyVal.data);(av[2]||[]).forEach(function(dvField){if(dvField.type!=='b')return;try{var dv=decodeMsg(dvField.data);(dv[1]||[]).forEach(function(f){if(f.type==='f32')EVOLVING_DATA[bName][eraKey].discrete[lv]=Math.round(readFloat32(f.data));else if(f.type==='v')EVOLVING_DATA[bName][eraKey].discrete[lv]=f.val;});}catch(e4){}});}catch(e3){}});
                                                }catch(e2){}
                                            });
                                            (change[2]||[]).forEach(function(fmAny){if(fmAny.type!=='b')return;try{var fmMsg=decodeMsg(fmAny.data);(fmMsg[2]||[]).forEach(function(fmPl){if(fmPl.type!=='b')return;try{var fp=decodeMsg(fmPl.data);(fp[1]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data))EVOLVING_DATA[bName][eraKey].formula=str(f.data);});}catch(e5){}});}catch(e5){}});
                                        }catch(e1){}
                                    });
                                }catch(e0){}
                            });
                        }catch(ex){}
                    });
                }catch(ex){}
            });
            log('Bâtiments évolutifs: '+Object.keys(EVOLVING_DATA).length+' parsés',Object.keys(EVOLVING_DATA).length>0?'#a8e6a3':'#f08080');
        }catch(e){log('Erreur parseEvolvingBuildings: '+e,'#f08080');}
    }

    function lookupEvolving(motif,era,level){
        var bData=null;
        for(var k in EVOLVING_DATA){if(motif.indexOf(k)!==-1){bData=EVOLVING_DATA[k];break;}}
        if(!bData){var stripped=motif.replace(/V\d+/,'');for(var k2 in EVOLVING_DATA){if(stripped.indexOf(k2)!==-1){bData=EVOLVING_DATA[k2];break;}}}
        if(!bData)return null;
        if(motif.indexOf('MinoanWatchtower')!==-1&&bData['DEFAULT']){
            var cult=level<=10?250+50*level:-250+100*level,rangeData=bData['range'],ray=0;
            if(rangeData){var rk=Object.keys(rangeData.discrete).map(Number).sort(function(a,b){return a-b;});for(var j=rk.length-1;j>=0;j--){if(rk[j]<=level){ray=rangeData.discrete[rk[j]];break;}}}
            return[cult,ray];
        }
        var eraOrder=['BronzeAge','MinoanEra','ClassicGreece','EarlyRome','RomanEmpire','ByzantineEra','AgeOfTheFranks','FeudalAge','IberianEra','KingdomOfSicily','HighMiddleAges','EarlyGothicEra','LateGothicEra'];
        var hasEraKeys=eraOrder.some(function(e){return!!bData[e];});
        var eraData=hasEraKeys?(bData[era]||function(){for(var i=eraOrder.length-1;i>=0;i--){if(bData[eraOrder[i]])return bData[eraOrder[i]];}return null;}()):bData['DEFAULT'];
        if(!eraData)return null;
        var cultVal=0,discrete=eraData.discrete,keys=Object.keys(discrete).map(Number).sort(function(a,b){return a-b;});
        var maxDL=keys.length>0?keys[keys.length-1]:0;
        if(level>maxDL){if(eraData.formula){var fv=applyFormula(eraData.formula,level);if(fv!==null)cultVal=Math.round(fv);}else if(keys.length>=2){var k1=keys[keys.length-2],k2=keys[keys.length-1];cultVal=Math.round(discrete[k2]+(discrete[k2]-discrete[k1])/(k2-k1)*(level-k2));}else if(keys.length===1)cultVal=discrete[keys[0]];}
        else{for(var i2=keys.length-1;i2>=0;i2--){if(keys[i2]<=level){cultVal=discrete[keys[i2]];break;}}}
        var rangeData2=bData['range'],ray2=0;
        if(rangeData2){var rk2=Object.keys(rangeData2.discrete).map(Number).sort(function(a,b){return a-b;});var maxRL=rk2.length>0?rk2[rk2.length-1]:0;if(level>maxRL&&rangeData2.formula){var rv=applyFormula(rangeData2.formula,level);if(rv!==null)ray2=Math.round(rv);}else{for(var j2=rk2.length-1;j2>=0;j2--){if(rk2[j2]<=level){ray2=rangeData2.discrete[rk2[j2]];break;}}}}
        return[cultVal,ray2];
    }

    function getInfo(fullName){
        var sv=STARTUP_CULTURE[fullName],cv=CATALOG[fullName];
        if(!cv){var base=fullName.replace(/_\d+$/,'');for(var k in CATALOG){if(k.replace(/_\d+$/,'')===base){cv=CATALOG[k];break;}}}
        var w=cv?cv.w:2,h=cv?cv.h:2;
        var culture=(sv&&sv.culture>0)?sv.culture:(cv?cv.culture:0);
        var range=cv?cv.range:0,t25=cv?(cv.t25||0):0,t50=cv?(cv.t50||0):0,t100=cv?(cv.t100||0):0;
        if(!t100){for(var k2 in DYNAMIC_THRESHOLDS){if(fullName.indexOf(k2)!==-1){t25=DYNAMIC_THRESHOLDS[k2][0];t50=DYNAMIC_THRESHOLDS[k2][1];t100=DYNAMIC_THRESHOLDS[k2][2];break;}}}
        return{w:w,h:h,culture:culture,range:range,t25:t25,t50:t50,t100:t100};
    }

    function computeCultureRay(name,level,era){
        var eraNom=(typeof era==='string'&&era.length>0)?era:'LateGothicEra';
        if(name.indexOf('CityHall')!==-1)return null;
        for(var fk in FIXED_CULTURAL){if(name.indexOf(fk)!==-1)return FIXED_CULTURAL[fk];}
        // Sites culturels DynamicAge : utiliser CULTURE_SITE_DATA
        for(var csk in CULTURE_SITE_DATA){
            if(name.indexOf(csk)!==-1){
                var csd=CULTURE_SITE_DATA[csk];
                var lvls=Object.keys(csd).map(Number).sort(function(a,b){return a-b;});
                var best=lvls[0];
                for(var li=0;li<lvls.length;li++){if(lvls[li]<=level)best=lvls[li];}
                return csd[best];
            }
        }
        for(var rfk in RANGE_FORMULAS){if(name.indexOf(rfk)!==-1){var ray=RANGE_FORMULAS[rfk][0][1];var res=lookupEvolving(rfk,eraNom,level);return res?[res[0],ray]:[0,ray];}}
        return lookupEvolving(name,eraNom,level)||[0,0];
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 4 : WASM / PARSING VILLE (repris de v5.4)
    // ═══════════════════════════════════════════════════════════════

    (function(){
        var origIS=WebAssembly.instantiateStreaming;
        WebAssembly.instantiateStreaming=function(){return origIS.apply(this,arguments).then(function(r){var inst=r.instance||r;if(inst&&inst.exports&&inst.exports.memory)window._rocWasmMem=inst.exports.memory;return r;});};
        var origI=WebAssembly.instantiate;
        WebAssembly.instantiate=function(){return origI.apply(this,arguments).then(function(r){var inst=r.instance||r;if(inst&&inst.exports&&inst.exports.memory&&!window._rocWasmMem)window._rocWasmMem=inst.exports.memory;return r;});};
    })();

    function findWasmMemory(){
        if(window._rocWasmMem)return true;
        var keys=Object.keys(window);
        for(var i=0;i<keys.length;i++){
            try{var v=window[keys[i]];if(!v||typeof v!=='object')continue;
                if(v instanceof WebAssembly.Memory){window._rocWasmMem=v;return true;}
                if(v.exports&&v.exports.memory instanceof WebAssembly.Memory){window._rocWasmMem=v.exports.memory;return true;}
                try{var sk=Object.keys(v);for(var j=0;j<Math.min(sk.length,50);j++){try{var v2=v[sk[j]];if(!v2||typeof v2!=='object')continue;if(v2 instanceof WebAssembly.Memory){window._rocWasmMem=v2;return true;}if(v2.exports&&v2.exports.memory instanceof WebAssembly.Memory){window._rocWasmMem=v2.exports.memory;return true;}}catch(e2){}}}catch(e3){}
            }catch(e){}
        }
        return false;
    }

    function autoFindCulture(pid){
        if(!window._rocWasmMem)findWasmMemory();
        if(!window._rocWasmMem)return;
        try{
            var mem32=new Uint32Array(window._rocWasmMem.buffer);
            if(window._rocCultureOffset){var cult=mem32[window._rocCultureOffset/4];if(cult>0&&cult<10000){if(pid&&window.cityMuseumData&&window.cityMuseumData[pid])window.cityMuseumData[pid].culture=cult;window._rocCultureValue=cult;log('Culture CityHall: '+cult,'#00ff88');return;}}
            var candidates=[],NULL32=4294967295;
            for(var i=3;i<mem32.length-2;i++){var v=mem32[i];if(v>=10&&v<=5000&&mem32[i-1]===NULL32&&mem32[i-2]===NULL32&&mem32[i+1]===NULL32&&mem32[i-3]===429)candidates.push({offset:i*4,value:v});}
            if(candidates.length===1){window._rocCultureOffset=candidates[0].offset;window._rocCultureValue=candidates[0].value;if(pid&&window.cityMuseumData&&window.cityMuseumData[pid])window.cityMuseumData[pid].culture=candidates[0].value;log('Culture CityHall: '+candidates[0].value,'#00ff88');}
        }catch(e){log('autoFindCulture err: '+e,'#f08080');}
    }

    function parseCityExtra(bytes){
        try{var root=decodeMsg(bytes);var cm=(root[1]||[])[0];if(!cm||cm.type!=='b')return;var city=decodeMsg(cm.data);var cd2=(city[2]||[])[0];if(!cd2||cd2.type!=='b')return;var cd=decodeMsg(cd2.data);var pid=0;(cd[1]||[]).forEach(function(f){if(f.type==='v')pid=f.val;});var range=1;(cd[8]||[]).forEach(function(cg){if(cg.type!=='b')return;try{var m=decodeMsg(cg.data);(m[1]||[]).forEach(function(f){if(f.type==='v'&&f.val>0)range=f.val;});}catch(e){}});var existing=window.cityMuseumData[pid]||{culture:0};window.cityMuseumData[pid]={range:range,culture:existing.culture};if(existing.culture===0)autoFindCulture(pid);}catch(e){}
    }

    function parseStartup(bytes){
        try{
            function searchInFields(fields,depth){
                if(depth>8)return;
                var name='';
                (fields[2]||[]).forEach(function(f){if(f.type==='b'&&isText(f.data)){var s=str(f.data);if(s.startsWith('Building_'))name=s;}});
                if(name){
                    (fields[21]||[]).forEach(function(f21item){if(f21item.type!=='b')return;try{var f21=decodeMsg(f21item.data);(f21[2]||[]).forEach(function(f){if(f.type==='v'&&f.val>0&&f.val<1000000){if(!STARTUP_CULTURE[name]||f.val>STARTUP_CULTURE[name].culture)STARTUP_CULTURE[name]={culture:f.val,range:STARTUP_CULTURE[name]?STARTUP_CULTURE[name].range:0};}});}catch(e){}});
                    [22,23,25,26].forEach(function(fn){(fields[fn]||[]).forEach(function(f){if(f.type==='v'&&f.val>0&&f.val<=10){if(STARTUP_CULTURE[name])STARTUP_CULTURE[name].range=f.val;else STARTUP_CULTURE[name]={culture:0,range:f.val};}});});
                }
                Object.keys(fields).forEach(function(fn){fields[fn].forEach(function(f){if(f.type==='b'&&f.data.length>5&&f.data.length<500000){try{searchInFields(decodeMsg(f.data),depth+1);}catch(e){}}});});
            }
            searchInFields(decodeMsg(bytes),0);
            log('Startup: '+Object.keys(STARTUP_CULTURE).length+' valeurs culture','#888');
        }catch(e){log('Erreur startup: '+e,'#f08080');}
        startupLoaded=true; lastStartupBytes=bytes;
        resetBuildings(); processCity(bytes); processPending();
        (function doWasmScan(attempt){
            if(window._rocWasmMem||findWasmMemory()){
                var pid=window._rocLastPlayerId;
                if(!window.cityMuseumData)window.cityMuseumData={};
                try{var so=localStorage.getItem('_rocCultureOffset');if(so)window._rocCultureOffset=parseInt(so);}catch(e){}
                if(pid>0&&!window.cityMuseumData[pid])window.cityMuseumData[pid]={range:0,culture:0};
                autoFindCulture(pid);
            }else if(attempt<5){setTimeout(function(){doWasmScan(attempt+1);},2000);}
        })(0);
    }

    function extractBuildings(bytes){
        var buildings=[],blocksLocked={};
        try{
            var root=decodeMsg(bytes);
            (root[1]||[]).forEach(function(f1item){
                if(f1item.type!=='b')return;
                var anyMsg=decodeMsg(f1item.data),typeUrl='';
                (anyMsg[1]||[]).forEach(function(f){if(f.type==='b')typeUrl=str(f.data);});
                var isOther=typeUrl.indexOf('OtherCityDTO')!==-1;
                var isCity=!isOther&&typeUrl.indexOf('CityDTO')!==-1;
                if(!isCity&&!isOther)return;
                (anyMsg[2]||[]).forEach(function(vf){
                    if(vf.type!=='b')return;
                    var cityDto=decodeMsg(vf.data),cityName='';
                    if(isCity){
                        (cityDto[2]||[]).forEach(function(f){if(f.type==='b')cityName=str(f.data);});
                        blocksLocked[cityName]=[];
                        (cityDto[6]||[]).forEach(function(blkItem){if(blkItem.type!=='b')return;var blk=decodeMsg(blkItem.data),col=null,row=null;(blk[2]||[]).forEach(function(f){if(f.type==='v')col=f.val;});(blk[3]||[]).forEach(function(f){if(f.type==='v')row=f.val>2147483647?f.val-4294967296:f.val;});if(col!==null&&row!==null&&row>=0)blocksLocked[cityName].push([col,row]);});
                        parseBuildingItems(cityDto[4]||[],cityName,buildings);
                    }else{
                        (cityDto[3]||[]).forEach(function(f){if(f.type==='b'&&!cityName)cityName=str(f.data);});
                        (cityDto[4]||[]).forEach(function(f){if(f.type==='b'&&!cityName)cityName=str(f.data);});
                        blocksLocked[cityName]=[];
                        (cityDto[6]||[]).forEach(function(blkItem){if(blkItem.type!=='b')return;try{var blk=decodeMsg(blkItem.data),col=null,row=null,w=0,h=0;(blk[2]||[]).forEach(function(f){if(f.type==='v')col=f.val;});(blk[3]||[]).forEach(function(f){if(f.type==='v')row=f.val>2147483647?f.val-4294967296:f.val;});(blk[6]||[]).forEach(function(f){if(f.type==='v')w=f.val;});(blk[7]||[]).forEach(function(f){if(f.type==='v')h=f.val;});if(col!==null&&row!==null&&row>=0&&w===4&&h===4)blocksLocked[cityName].push([col,row]);}catch(e){}});
                        parseBuildingItems(cityDto[5]||[],cityName,buildings);
                    }
                });
            });
        }catch(e){}
        return{buildings:buildings,blocksLocked:blocksLocked};
    }

    function parseBuildingItems(items,cityName,buildings){
        items.forEach(function(bldItem){
            if(bldItem.type!=='b')return;
            var bld=decodeMsg(bldItem.data),name='',col=null,row=null,level=1,rotated=0,era='',playerId=0,cityDtoCulture=0,cityDtoRange=0;
            (bld[2]||[]).forEach(function(f){if(f.type==='b')name=str(f.data);});
            (bld[4]||[]).forEach(function(f){if(f.type==='v')col=f.val>2147483647?f.val-4294967296:f.val;});
            (bld[5]||[]).forEach(function(f){if(f.type==='v')row=f.val>2147483647?f.val-4294967296:f.val;});
            (bld[13]||[]).forEach(function(f){if(f.type==='v')playerId=f.val;});
            if(playerId>0)window._rocLastPlayerId=playerId;
            (bld[17]||[]).forEach(function(f){if(f.type==='b')era=str(f.data);});
            (bld[18]||[]).forEach(function(f){if(f.type==='v')level=f.val;});
            (bld[9]||[]).forEach(function(f){if(f.type==='v')rotated=f.val;});
            (bld[21]||[]).forEach(function(f21item){if(f21item.type!=='b')return;try{var f21=decodeMsg(f21item.data);(f21[2]||[]).forEach(function(f){if(f.type==='v'&&f.val>0&&f.val<1000000)cityDtoCulture=f.val;});(f21[3]||[]).forEach(function(f){if(f.type==='v'&&f.val>0&&f.val<=10)cityDtoRange=f.val;});}catch(e){}});
            if(name&&col!==null&&row!==null)buildings.push({city:cityName,name:name,col:col,row:row,level:level,era:era,rotated:rotated,cityDtoCulture:cityDtoCulture,cityDtoRange:cityDtoRange,playerId:playerId});
        });
    }

    function resetBuildings(){allBuildings=[];cities={};allLockedBlocks={};lockedBlocksKeys={};log('🔄 Nouvelle ville — données réinitialisées','#f0c040');}

    function processCity(bytes){
        if(!startupLoaded){pendingCityBytes.push(bytes);return;}
        var extracted=extractBuildings(bytes);
        var buildings=extracted.buildings||[],lockedBlocks=extracted.blocksLocked||{};
        Object.keys(lockedBlocks).forEach(function(city){if(!allLockedBlocks[city])allLockedBlocks[city]=[];lockedBlocks[city].forEach(function(b){var bkey=city+'|'+b[0]+'|'+b[1];if(!lockedBlocksKeys[bkey]){lockedBlocksKeys[bkey]=true;allLockedBlocks[city].push(b);}});});
        var nbLocked=Object.values(lockedBlocks).reduce(function(s,a){return s+a.length;},0);
        log('✓ '+buildings.length+' bâtiments, '+nbLocked+' blocs achetés vides','#a8e6a3');
        buildings.forEach(function(b){var key=b.city+'|'+b.name+'|'+b.col+'|'+b.row;if(!cities[key]){cities[key]=true;allBuildings.push(b);}});
        updatePanel();
    }
    function processPending(){var p=pendingCityBytes.slice();pendingCityBytes=[];p.forEach(function(b){processCity(b);});}

    // Interception réseau
    var origOpen=XMLHttpRequest.prototype.open,origSend=XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open=function(m,u){this._roc_url=u;return origOpen.apply(this,arguments);};
    XMLHttpRequest.prototype.send=function(){
        var self=this;
        this.addEventListener('load',function(){
            var url=String(self._roc_url||'');
            if(!(self.response instanceof ArrayBuffer))return;
            var bytes=new Uint8Array(self.response);
            if(url.indexOf('/game/startup')!==-1)parseStartup(bytes);
            else if(url.indexOf('/game/visit-city')!==-1){resetBuildings();processCity(bytes);}
            else if(url.indexOf('/game/city')!==-1)parseCityExtra(bytes);
        });
        return origSend.apply(this,arguments);
    };
    var origFetch=window.fetch;
    window.fetch=function(url,opts){
        var urlStr=String(url);
        return origFetch.apply(this,arguments).then(function(response){
            var isSt=urlStr.indexOf('/game/startup')!==-1,isV=urlStr.indexOf('/game/visit-city')!==-1,isC=urlStr.indexOf('/game/city')!==-1;
            if(isSt||isV||isC){response.clone().arrayBuffer().then(function(buf){var bytes=new Uint8Array(buf);if(isSt)parseStartup(bytes);else if(isV){resetBuildings();processCity(bytes);}else parseCityExtra(bytes);}).catch(function(){});}
            return response;
        });
    };

    // ═══════════════════════════════════════════════════════════════
    // SECTION 5 : HELPERS NOMMAGE & CATÉGORIE
    // ═══════════════════════════════════════════════════════════════

    function csvBuildingCategory(nom){
        var n=nom.toUpperCase();
        if(n.indexOf('BARRACKS')!==-1)return 'Barracks';
        if(n.indexOf('FARM')!==-1||n.indexOf('CAMELFARMED')!==-1||n.indexOf('IRRIGATION')!==-1||n.indexOf('CAMEL')!==-1)return 'Farm';
        if(n.indexOf('HOME')!==-1)return 'Home';
        if(n.indexOf('WORKSHOP')!==-1||n.indexOf('SMITHY')!==-1||n.indexOf('FORGE')!==-1||n.indexOf('MERCHANT')!==-1||n.indexOf('CARPENTER')!==-1||n.indexOf('POTTERY')!==-1||n.indexOf('WEAVER')!==-1||n.indexOf('MASON')!==-1||n.indexOf('TANNERY')!==-1||n.indexOf('BREWERY')!==-1||n.indexOf('MILL')!==-1||n.indexOf('BAKERY')!==-1||n.indexOf('GLASSWORKS')!==-1||n.indexOf('JEWELLER')!==-1||n.indexOf('SCRIPTORIUM')!==-1||n.indexOf('ARMORY')!==-1||n.indexOf('SADDLER')!==-1)return 'Workshop';
        return null;
    }

    function cleanBuildingName(nom){
        var s=nom;if(s.startsWith('Building_'))s=s.slice(9);
        var parts=s.split('_'),level='';
        if(parts.length>0&&/^\d+$/.test(parts[parts.length-1])){level=' Lv'+parts[parts.length-1];parts=parts.slice(0,-1);}
        if(parts.length>1&&!parts[0].startsWith('Event'))parts=parts.slice(1);
        return parts.join(' ')+level;
    }

    function isFixedObstacle(name){
    var n=name.toUpperCase();
    // Obstacles fixes (non déplaçables)
    if(n.indexOf('CONNECTION')!==-1||n.indexOf('PIER')!==-1) return true;
    // Bâtiments à ignorer (harbor/port)
    if(n.indexOf('WAREHOUSE')!==-1) return true;  // Entrepôt
    if(n.indexOf('SAILOR')!==-1) return true;      // Maison de marins
    if(n.indexOf('SHIPYARD')!==-1) return true;    // Chantier naval
    if(n.indexOf('HARBOR')!==-1) return true;      // Harbor général
    return false;
}

    function pad(n){return n<10?'0'+n:''+n;}
    function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
    function log(msg,color){var entry={msg:msg,color:color||'#aaa'};logs.push(entry);fullLog.push(msg);if(logs.length>30)logs.shift();updatePanel();}

    // ═══════════════════════════════════════════════════════════════
    // SECTION 6 : CONVERSION ALLBUILDINGS → DONNÉES OPTIMISEUR
    // ═══════════════════════════════════════════════════════════════

    function detectBuildingType(name, t100, culture, range) {
        if(t100>0)return 'Producteur';
        var cat=csvBuildingCategory(name);
        if(cat==='Farm'||cat==='Home'||cat==='Barracks')return 'Producteur';
        if(range>0&&(name.indexOf('CultureSite')!==-1||name.indexOf('Evolving')!==-1||name.indexOf('Collectable')!==-1))return 'Culturel';
        if(culture>0&&range>0)return 'Culturel';
        return 'Neutre';
    }

    function buildOptimData(prio_par_type, boost100_required_set) {
        prio_par_type = prio_par_type || {};
        boost100_required_set = boost100_required_set || new Set();
        var cityName = allBuildings.length > 0 ? allBuildings[0].city : '';

        // Enrichir chaque bâtiment
        var enriched = [];
        allBuildings.forEach(function(b) {
            var info = getInfo(b.name);
            var level = b.level || 1;
            var era = b.era || '';
            var culture = info.culture, range = info.range;
            var isCT = b.name.indexOf('Evolving')!==-1||b.name.indexOf('Collectable')!==-1||b.name.indexOf('DynamicAge')!==-1||b.name.indexOf('CityHall')!==-1;

            if(b.name.indexOf('CityHall')!==-1){
                var md=window.cityMuseumData&&window.cityMuseumData[b.playerId];
                if(md&&md.range>0)range=md.range;
                // Culture CityHall: STARTUP_CULTURE > WASM > cityDtoCulture
                var _sc=STARTUP_CULTURE[b.name];
                if(_sc&&_sc.culture>0){
                    culture=_sc.culture;
                    log('CityHall STARTUP_CULTURE: '+culture,'#a8e6a3');
                } else if(window._rocCultureValue>0){
                    culture=window._rocCultureValue;
                    log('CityHall _rocCultureValue: '+culture,'#f0c040');
                } else if(window._rocWasmMem&&window._rocCultureOffset){
                    var m32c=new Uint32Array(window._rocWasmMem.buffer);
                    var cv0=m32c[window._rocCultureOffset/4];
                    if(cv0>0&&cv0<10000){culture=cv0;window._rocCultureValue=cv0;
                        log('CityHall WASM: '+culture,'#f0c040');}
                }
                if(culture===0&&b.cityDtoCulture>0){culture=b.cityDtoCulture;
                    log('CityHall cityDtoCulture: '+culture,'#f08080');}
                if(culture===0&&md&&md.culture>0)culture=md.culture;
            }else if(isCT){
                // Toujours recalculer la culture avec le niveau réel du bâtiment
                var cr=computeCultureRay(b.name,level,era);
                if(cr&&cr[0]>0)culture=cr[0];
                if(cr&&cr[1]>0)range=cr[1];
            }
            // Version multi-niveaux (fallback si computeCultureRay échoue)
            if(culture===0){
                var vn=b.name.replace(/_\d+$/,'_'+level);
                if(vn!==b.name&&CATALOG[vn]){if(CATALOG[vn].culture>0)culture=CATALOG[vn].culture;if(CATALOG[vn].range>0)range=CATALOG[vn].range;}
            }
            if(culture===0&&b.cityDtoCulture>0)culture=b.cityDtoCulture;

            var w=b.rotated?info.h:info.w, h=b.rotated?info.w:info.h;
            enriched.push({name:b.name,col:b.col,row:b.row,level:level,era:era,rotated:b.rotated,playerId:b.playerId,w:w,h:h,culture:culture,range:range,t25:info.t25,t50:info.t50,t100:info.t100,nom_fr:getLocaName(b.name)||''});
        });

        // Cases fixes (obstacles)
        var fixedSet = new Set();
        enriched.forEach(function(b){if(isFixedObstacle(b.name)){for(var dr=0;dr<b.h;dr++)for(var dc=0;dc<b.w;dc++)fixedSet.add((b.row+dr)+'|'+(b.col+dc));}});

        // Terrain
        var validInterior = new Set();
        var lockedBlocks = allLockedBlocks[cityName] || [];
        enriched.forEach(function(b){if(isFixedObstacle(b.name))return;for(var dr=0;dr<b.h;dr++)for(var dc=0;dc<b.w;dc++){var k=(b.row+dr)+'|'+(b.col+dc);if(!fixedSet.has(k))validInterior.add(k);}});
        lockedBlocks.forEach(function(lb){for(var dr=0;dr<4;dr++)for(var dc=0;dc<4;dc++){var k=(lb[1]+dr)+'|'+(lb[0]+dc);if(!fixedSet.has(k))validInterior.add(k);}});
        if(validInterior.size===0)return null;

        var allR=[],allC=[];
        validInterior.forEach(function(k){var p=k.split('|');allR.push(parseInt(p[0]));allC.push(parseInt(p[1]));});
        var lig_min=Math.min.apply(null,allR),lig_max=Math.max.apply(null,allR);
        var col_min=Math.min.apply(null,allC),col_max=Math.max.apply(null,allC);
        var margin=1;
        var gLM=lig_min-margin,gCM=col_min-margin;
        var max_r=(lig_max+margin)-gLM+1,max_c=(col_max+margin)-gCM+1;

        var terrain_grid=[];
        for(var r=0;r<max_r;r++){var row_=[];for(var c=0;c<max_c;c++)row_.push('X');terrain_grid.push(row_);}
        validInterior.forEach(function(k){var p=k.split('|');var rg=parseInt(p[0])-gLM,cg=parseInt(p[1])-gCM;if(rg>=0&&rg<max_r&&cg>=0&&cg<max_c)terrain_grid[rg][cg]=null;});

        // Catalogue
        var catalog_rows = {}, realEnriched = enriched.filter(function(b){return !isFixedObstacle(b.name);});
        realEnriched.forEach(function(b){
            var key=b.name;
            if(!catalog_rows[key]){
                var btype=detectBuildingType(b.name,b.t100,b.culture,b.range);
                var prio=btype==='Producteur'?(prio_par_type[key]||0):0;
                catalog_rows[key]={nom:key,nom_fr:b.nom_fr,longueur:b.w,largeur:b.h,nombre:0,type:btype,culture:b.culture,rayonnement:b.range,boost25:b.t25,boost50:b.t50,boost100:b.t100,production:'',quantite:0,priorite:prio,placement:'Obligatoire',boost100_required:boost100_required_set.has(key)};
            }
            catalog_rows[key].nombre++;
        });
        var buildings_def = Object.values(catalog_rows);

        // Bâtiments placés
        var placed = [];
        realEnriched.forEach(function(b){
            var bdef=catalog_rows[b.name];
            placed.push(Object.assign({},bdef,{nom:b.name,nom_fr:b.nom_fr,r:b.row-gLM,c:b.col-gCM,rows:b.h,cols:b.w,culture:b.culture,rayonnement:b.range,era:b.era,level:b.level||1}));
        });

        return{terrain_grid:terrain_grid,max_r:max_r,max_c:max_c,placed:placed,buildings_def:buildings_def,grid_lig_min:gLM,grid_col_min:gCM};
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 7 : MÉCANIQUE CULTURE / BOOST / SCORE
    // ═══════════════════════════════════════════════════════════════

    function cultureReceived(producer, culturels) {
        var pr=producer.r,pc_=producer.c,prows=producer.rows,pcols=producer.cols,total=0;
        for(var ci=0;ci<culturels.length;ci++){
            var cult=culturels[ci],ray=cult.rayonnement||0;if(ray===0)continue;
            var r0=cult.r,c0=cult.c,r1=r0+cult.rows-1,c1=c0+cult.cols-1;
            var found=false;
            outer:for(var dr=0;dr<prows;dr++){for(var dc=0;dc<pcols;dc++){var rr=pr+dr,cc=pc_+dc;if(r0-ray<=rr&&rr<=r1+ray&&c0-ray<=cc&&cc<=c1+ray&&!(r0<=rr&&rr<=r1&&c0<=cc&&cc<=c1)){found=true;break outer;}}}
            if(found)total+=cult.culture;
        }
        return total;
    }

    function boostLevel(culture,b){
        if(b.type!=='Producteur')return 0;
        if(b.boost100&&culture>=b.boost100)return 100;
        if(b.boost50&&culture>=b.boost50)return 50;
        if(b.boost25&&culture>=b.boost25)return 25;
        return 0;
    }

    function scorePlacement(placed){
        var culturels=[],total=0;
        for(var i=0;i<placed.length;i++){if(placed[i].type==='Culturel')culturels.push(placed[i]);}
        for(var i=0;i<placed.length;i++){
            var b=placed[i];
            if(b.type!=='Producteur')continue;
            var cult=cultureReceived(b,culturels);
            var bl=boostLevel(cult,b);
            // Score normal (priorite>0)
            if(b.priorite>0)total+=bl*b.priorite;
            // Pénalité boost100_required
            if(b.boost100_required&&bl<100)total-=10000;
            // Pénalité no_reduction : s'applique même si priorite=0
            if(b.boost_min!==undefined&&bl<b.boost_min)total-=1000000;
        }
        return total;
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 8 : MAKESGRID & CANPLACE
    // ═══════════════════════════════════════════════════════════════

    function makeXGrid(terrain_grid,max_r,max_c){
        var exterior=new Uint8Array(max_r*max_c),queue=[],W=max_c;
        function seed(r,c){var idx=r*W+c;if(!exterior[idx]&&terrain_grid[r][c]!=='X'){exterior[idx]=1;queue.push(r,c);}}
        for(var r=0;r<max_r;r++){seed(r,0);seed(r,max_c-1);}
        for(var c=0;c<max_c;c++){seed(0,c);seed(max_r-1,c);}
        var qi=0;
        while(qi<queue.length){var qr=queue[qi++],qc=queue[qi++];var nb=[[qr-1,qc],[qr+1,qc],[qr,qc-1],[qr,qc+1]];for(var n=0;n<4;n++){var nr=nb[n][0],nc=nb[n][1];if(nr<0||nr>=max_r||nc<0||nc>=max_c)continue;var idx2=nr*W+nc;if(!exterior[idx2]&&terrain_grid[nr][nc]!=='X'){exterior[idx2]=1;queue.push(nr,nc);}}}
        var xg=[];for(var r2=0;r2<max_r;r2++){var row=new Uint8Array(max_c);for(var c2=0;c2<max_c;c2++)row[c2]=(terrain_grid[r2][c2]==='X'||exterior[r2*W+c2])?1:0;xg.push(row);}
        return xg;
    }

    function buildOcc(placed){
        var occ=new Set();
        for(var i=0;i<placed.length;i++){var b=placed[i];for(var dr=0;dr<b.rows;dr++)for(var dc=0;dc<b.cols;dc++)occ.add((b.r+dr)+'|'+(b.c+dc));}
        return occ;
    }

    function canPlace(r,c,rows,cols,x_grid,occ,max_r,max_c){
        if(r<0||c<0||r+rows>max_r||c+cols>max_c)return false;
        for(var dr=0;dr<rows;dr++)for(var dc=0;dc<cols;dc++){if(x_grid[r+dr][c+dc]||occ.has((r+dr)+'|'+(c+dc)))return false;}
        return true;
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 9 : PLACEMENT DES BÂTIMENTS MANQUANTS
    // ═══════════════════════════════════════════════════════════════

    function placeMissingBuildings(placed,buildings_def,terrain_grid,max_r,max_c){
        var placedCounts={};
        placed.forEach(function(b){placedCounts[b.nom]=(placedCounts[b.nom]||0)+1;});
        var toPlace=[];
        buildings_def.forEach(function(bdef){
            var needed=bdef.nombre-(placedCounts[bdef.nom]||0);
            for(var i=0;i<needed;i++)toPlace.push(Object.assign({},bdef,{rows:bdef.largeur,cols:bdef.longueur}));
        });
        if(toPlace.length===0)return placed.map(function(b){return Object.assign({},b);});

        var x_grid=makeXGrid(terrain_grid,max_r,max_c);

        // BFS dist to X border
        var dist_to_x=[];
        for(var r=0;r<max_r;r++){var row_=[];for(var c=0;c<max_c;c++)row_.push(999);dist_to_x.push(row_);}
        var bfsQ=[],bfsQi=0;
        for(var r=0;r<max_r;r++)for(var c=0;c<max_c;c++){if(terrain_grid[r][c]==='X'){dist_to_x[r][c]=0;bfsQ.push(r,c);}}
        while(bfsQi<bfsQ.length){var qr=bfsQ[bfsQi++],qc=bfsQ[bfsQi++];var nb=[[qr-1,qc],[qr+1,qc],[qr,qc-1],[qr,qc+1]];for(var ni=0;ni<4;ni++){var nr=nb[ni][0],nc=nb[ni][1];if(nr<0||nr>=max_r||nc<0||nc>=max_c)continue;if(dist_to_x[nr][nc]===999){dist_to_x[nr][nc]=dist_to_x[qr][qc]+1;bfsQ.push(nr,nc);}}}

        // Cases intérieures triées par distance au bord (pour les Neutres)
        var innerCells=[];
        for(var r2=0;r2<max_r;r2++)for(var c2=0;c2<max_c;c2++){if(!x_grid[r2][c2])innerCells.push([r2,c2,dist_to_x[r2][c2]]);}
        innerCells.sort(function(a,b){return a[2]-b[2];});

        // Tri: Producteurs > Culturels > Neutres, puis grands en premier
        var typePrio={Producteur:0,Culturel:1,Neutre:2};
        toPlace.sort(function(a,b){var ta=typePrio[a.type]||2,tb=typePrio[b.type]||2;return ta!==tb?ta-tb:(b.rows*b.cols)-(a.rows*a.cols);});

        var result=placed.map(function(b){return Object.assign({},b);});
        var nFail=0;

        toPlace.forEach(function(b){
            var occ=buildOcc(result);
            var orientations=[[b.rows,b.cols]];
            if(b.rows!==b.cols)orientations.push([b.cols,b.rows]);
            var placed_ok=false;

            if(b.type==='Neutre'){
                // Essayer depuis les bords
                outer_n:for(var oi=0;oi<orientations.length;oi++){var rows=orientations[oi][0],cols=orientations[oi][1];for(var ci=0;ci<innerCells.length;ci++){var r=innerCells[ci][0],c=innerCells[ci][1];if(canPlace(r,c,rows,cols,x_grid,occ,max_r,max_c)){result.push(Object.assign({},b,{r:r,c:c,rows:rows,cols:cols}));placed_ok=true;break outer_n;}}}
            }else{
                // Scan raster normal
                outer_p:for(var oi2=0;oi2<orientations.length;oi2++){var rows2=orientations[oi2][0],cols2=orientations[oi2][1];for(var r2=0;r2<max_r;r2++){for(var c2=0;c2<max_c;c2++){if(canPlace(r2,c2,rows2,cols2,x_grid,occ,max_r,max_c)){result.push(Object.assign({},b,{r:r2,c:c2,rows:rows2,cols:cols2}));placed_ok=true;break outer_p;}}}}
            }
            if(!placed_ok)nFail++;
        });

        if(nFail>0)log('⚠️ '+nFail+' bâtiment(s) non placés (terrain plein)','#f08080');
        return result;
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 10 : OPTIMISEUR (greedy + swaps de 3 + SA)
    // ═══════════════════════════════════════════════════════════════

    function bestPositionFor(b,placed,x_grid,max_r,max_c){
        if(b.type==='Neutre')return null;
        var occSet=buildOcc(placed.filter(function(p){return p!==b;}));
        var origScore=scorePlacement(placed);
        var bestScore=origScore,bestPos=null;
        var origR=b.r,origC=b.c,origRows=b.rows,origCols=b.cols;
        var orientations=[[b.rows,b.cols]];if(b.rows!==b.cols)orientations.push([b.cols,b.rows]);
        for(var oi=0;oi<orientations.length;oi++){
            var rows=orientations[oi][0],cols=orientations[oi][1];
            for(var r=0;r<max_r;r++){for(var c=0;c<max_c;c++){
                if(r===origR&&c===origC&&rows===origRows&&cols===origCols)continue;
                if(!canPlace(r,c,rows,cols,x_grid,occSet,max_r,max_c))continue;
                b.r=r;b.c=c;b.rows=rows;b.cols=cols;
                var s=scorePlacement(placed);
                if(s>bestScore){bestScore=s;bestPos=[r,c,rows,cols];}
                b.r=origR;b.c=origC;b.rows=origRows;b.cols=origCols;
            }}
        }
        return bestPos;
    }

    function greedyPass(placed,x_grid,max_r,max_c){
        var improved=false;
        var sorted=placed.filter(function(b){return b.type!=='Neutre';}).sort(function(a,b){return b.priorite-a.priorite;});
        for(var i=0;i<sorted.length;i++){var b=sorted[i];var pos=bestPositionFor(b,placed,x_grid,max_r,max_c);if(pos){b.r=pos[0];b.c=pos[1];b.rows=pos[2];b.cols=pos[3];improved=true;}}
        return improved;
    }

    function trySwap3(a,b,c,placed,x_grid,max_r,max_c,scoreBefore){
        var sa=[a.r,a.c,a.rows,a.cols],sb=[b.r,b.c,b.rows,b.cols],sc=[c.r,c.c,c.rows,c.cols];
        var occBase=buildOcc(placed.filter(function(p){return p!==a&&p!==b&&p!==c;}));

        function noOverlap(p1,p2,p3){
            for(var dr=0;dr<p1[2];dr++)for(var dc=0;dc<p1[3];dc++){var k=(p1[0]+dr)+'|'+(p1[1]+dc);for(var dr2=0;dr2<p2[2];dr2++)for(var dc2=0;dc2<p2[3];dc2++)if(k===(p2[0]+dr2)+'|'+(p2[1]+dc2))return false;for(var dr3=0;dr3<p3[2];dr3++)for(var dc3=0;dc3<p3[3];dc3++)if(k===(p3[0]+dr3)+'|'+(p3[1]+dc3))return false;}
            for(var dr4=0;dr4<p2[2];dr4++)for(var dc4=0;dc4<p2[3];dc4++){var k2=(p2[0]+dr4)+'|'+(p2[1]+dc4);for(var dr5=0;dr5<p3[2];dr5++)for(var dc5=0;dc5<p3[3];dc5++)if(k2===(p3[0]+dr5)+'|'+(p3[1]+dc5))return false;}
            return true;
        }

        var bestDelta=0,bestCfg=null;
        var oriA=[[a.rows,a.cols]],oriB=[[b.rows,b.cols]],oriC=[[c.rows,c.cols]];
        if(a.rows!==a.cols)oriA.push([a.cols,a.rows]);
        if(b.rows!==b.cols)oriB.push([b.cols,b.rows]);
        if(c.rows!==c.cols)oriC.push([c.cols,c.rows]);

        for(var i=0;i<oriA.length;i++)for(var j=0;j<oriB.length;j++)for(var k2=0;k2<oriC.length;k2++){
            // a→pos(b), b→pos(c), c→pos(a)
            var pa=[sb[0],sb[1],oriA[i][0],oriA[i][1]];
            var pb=[sc[0],sc[1],oriB[j][0],oriB[j][1]];
            var pc=[sa[0],sa[1],oriC[k2][0],oriC[k2][1]];
            if(!noOverlap(pa,pb,pc))continue;
            if(!canPlace(pa[0],pa[1],pa[2],pa[3],x_grid,occBase,max_r,max_c))continue;
            if(!canPlace(pb[0],pb[1],pb[2],pb[3],x_grid,occBase,max_r,max_c))continue;
            if(!canPlace(pc[0],pc[1],pc[2],pc[3],x_grid,occBase,max_r,max_c))continue;
            a.r=pa[0];a.c=pa[1];a.rows=pa[2];a.cols=pa[3];
            b.r=pb[0];b.c=pb[1];b.rows=pb[2];b.cols=pb[3];
            c.r=pc[0];c.c=pc[1];c.rows=pc[2];c.cols=pc[3];
            var s=scorePlacement(placed),delta=s-scoreBefore;
            if(delta>bestDelta){bestDelta=delta;bestCfg=[pa,pb,pc];}
            a.r=sa[0];a.c=sa[1];a.rows=sa[2];a.cols=sa[3];
            b.r=sb[0];b.c=sb[1];b.rows=sb[2];b.cols=sb[3];
            c.r=sc[0];c.c=sc[1];c.rows=sc[2];c.cols=sc[3];
        }
        if(bestCfg){
            a.r=bestCfg[0][0];a.c=bestCfg[0][1];a.rows=bestCfg[0][2];a.cols=bestCfg[0][3];
            b.r=bestCfg[1][0];b.c=bestCfg[1][1];b.rows=bestCfg[1][2];b.cols=bestCfg[1][3];
            c.r=bestCfg[2][0];c.c=bestCfg[2][1];c.rows=bestCfg[2][2];c.cols=bestCfg[2][3];
            return true;
        }
        return false;
    }

    function optimizeMultiswap(placed,terrain_grid,max_r,max_c,progress_cb,time_budget_ms,mode){
        mode=mode||'priority'; time_budget_ms=time_budget_ms||60000;
        var x_grid=makeXGrid(terrain_grid,max_r,max_c);
        placed=placed.map(function(b){return Object.assign({},b);});
        var t_start=Date.now();
        function elapsed(){return Date.now()-t_start;}
        function prog(f){if(progress_cb)progress_cb(Math.min(f,0.98));}

        var T_GREEDY=time_budget_ms*0.10,T_SW3=time_budget_ms*0.55,T_SA=time_budget_ms*0.35;
        prog(0);

        // Phase 0 : greedy
        for(var gi=0;gi<8&&elapsed()<T_GREEDY;gi++){if(!greedyPass(placed,x_grid,max_r,max_c))break;}
        prog(0.10);

        // Phase 1 : swaps de 3
        var cands=placed.filter(function(b){return b.type==='Culturel'||b.type==='Producteur';});
        var sw3Improved=true;
        while(sw3Improved&&elapsed()<T_GREEDY+T_SW3){
            sw3Improved=false;
            var scoreBefore=scorePlacement(placed);
            for(var i=0;i<cands.length&&elapsed()<T_GREEDY+T_SW3;i++){
                for(var j=i+1;j<cands.length&&elapsed()<T_GREEDY+T_SW3;j++){
                    for(var k=j+1;k<cands.length&&elapsed()<T_GREEDY+T_SW3;k++){
                        if(trySwap3(cands[i],cands[j],cands[k],placed,x_grid,max_r,max_c,scoreBefore)){scoreBefore=scorePlacement(placed);sw3Improved=true;}
                    }
                }
            }
            if(sw3Improved)greedyPass(placed,x_grid,max_r,max_c);
            prog(0.10+0.55*Math.min(elapsed()/(T_GREEDY+T_SW3),1));
        }
        prog(0.65);

        // Phase 2 : SA léger
        var saScore=scorePlacement(placed);
        var saBest=placed.map(function(b){return Object.assign({},b);}),saBestScore=saScore;
        var T_sa=500,alpha_sa=0.995,T_min_sa=0.5,saIter=0;
        while(T_sa>T_min_sa&&elapsed()<time_budget_ms-500){
            if(cands.length<2){break;}
            var i1=Math.floor(Math.random()*cands.length),i2=Math.floor(Math.random()*cands.length);
            if(i1===i2){T_sa*=alpha_sa;continue;}
            var b1=cands[i1],b2=cands[i2];
            var s1=[b1.r,b1.c,b1.rows,b1.cols],s2=[b2.r,b2.c,b2.rows,b2.cols];
            var occSA=buildOcc(placed.filter(function(p){return p!==b1&&p!==b2;}));
            var swapped=false;
            outer_sa:for(var oa=0;oa<2;oa++){for(var ob=0;ob<2;ob++){
                var r1=oa===0?b1.rows:b1.cols,c1=oa===0?b1.cols:b1.rows;
                var r2=ob===0?b2.rows:b2.cols,c2=ob===0?b2.cols:b2.rows;
                if(!canPlace(s2[0],s2[1],r1,c1,x_grid,occSA,max_r,max_c))continue;
                // Check no overlap between the two new positions
                var ov=false;
                for(var dr_=0;dr_<r1;dr_++)for(var dc_=0;dc_<c1;dc_++){var k_=(s2[0]+dr_)+'|'+(s2[1]+dc_);for(var dr2_=0;dr2_<r2;dr2_++)for(var dc2_=0;dc2_<c2;dc2_++)if(k_===(s1[0]+dr2_)+'|'+(s1[1]+dc2_)){ov=true;break;}}
                if(ov)continue;
                if(!canPlace(s1[0],s1[1],r2,c2,x_grid,occSA,max_r,max_c))continue;
                b1.r=s2[0];b1.c=s2[1];b1.rows=r1;b1.cols=c1;
                b2.r=s1[0];b2.c=s1[1];b2.rows=r2;b2.cols=c2;
                swapped=true;break outer_sa;
            }}
            if(!swapped){T_sa*=alpha_sa;saIter++;continue;}
            var ns=scorePlacement(placed),dl=ns-saScore;
            if(dl>0||Math.random()<Math.exp(dl/T_sa)){saScore=ns;if(ns>saBestScore){saBestScore=ns;saBest=placed.map(function(b){return Object.assign({},b);});}}
            else{b1.r=s1[0];b1.c=s1[1];b1.rows=s1[2];b1.cols=s1[3];b2.r=s2[0];b2.c=s2[1];b2.rows=s2[2];b2.cols=s2[3];}
            T_sa*=alpha_sa;saIter++;
            if(saIter%500===0)prog(0.65+0.33*Math.min(elapsed()/time_budget_ms,1));
        }
        // Restaurer meilleur SA
        for(var bi=0;bi<placed.length&&bi<saBest.length;bi++){placed[bi].r=saBest[bi].r;placed[bi].c=saBest[bi].c;placed[bi].rows=saBest[bi].rows;placed[bi].cols=saBest[bi].cols;}
        greedyPass(placed,x_grid,max_r,max_c);
        prog(1.0);
        return placed;
    }

    // ═══════════════════════════════════════════════════════════════
    // SECTION 11 : GÉNÉRATION EXCEL (SheetJS)
    // ═══════════════════════════════════════════════════════════════

    // Couleurs ARGB
    var C_ORANGE='FFA500',C_GREEN='90EE90',C_GRAY='D3D3D3',C_BLUE='4472C4';
    var C_BORDX='808080',C_BOOST0='FFD7D7',C_BOOST25='FFF2CC',C_BOOST50='D9EAD3',C_BOOST100='93C47D';
    var C_MOVE='CFE2F3',C_RED='CC0000',C_DARKGREEN='006400';

    function colLetter(n){var s='';while(n>0){var r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}

    function coordExcel(r0,c0,max_r,max_c){
        var newCol=max_r-r0,newRow=max_c-c0;
        if(newCol<1||newRow<1)return'';
        return colLetter(newCol)+newRow;
    }

    function makeCell(v,opts){
        opts=opts||{};
        var cell={v:v,t:typeof v==='number'?'n':'s'};
        var style={};
        if(opts.bold||opts.color||opts.bg||opts.align!==undefined||opts.wrap){
            var font={};
            if(opts.bold)font.bold=true;
            if(opts.color)font.color={rgb:opts.color};
            if(Object.keys(font).length)style.font=font;
            if(opts.bg)style.fill={fgColor:{rgb:opts.bg},patternType:'solid',bgColor:{indexed:64}};
            style.alignment={horizontal:opts.align||'center',vertical:'center',wrapText:!!opts.wrap};
            style.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
            cell.s=style;
        }
        return cell;
    }

    function headerCell(text){return makeCell(text,{bold:true,color:'FFFFFF',bg:C_BLUE,align:'center'});}

    function displayName(b){return b.nom_fr||cleanBuildingName(b.nom||'');}

    function buildExcelOutput(optimized, originalPlaced, terrain_grid, max_r, max_c, buildings_def, boost100_req, prio_par_type, optim_mode, protected_cats){
        boost100_req = boost100_req || new Set();
        var wb = XLSX.utils.book_new();

        // Styles cellules (xlsx-js-style / SheetJS Pro compatible)
        var S = {
            hdr:  function(){ return {fill:{patternType:'solid',fgColor:{rgb:'1F4E79'}},font:{bold:true,color:{rgb:'FFFFFF'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            cult: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'FFA500'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            prod: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'90EE90'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            neut: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'D3D3D3'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            miss: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'FFD7D7'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            move: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'CFE2F3'}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}}; },
            bx:   function(){ return {fill:{patternType:'solid',fgColor:{rgb:'808080'}},font:{bold:true,color:{rgb:'FFFFFF'}},alignment:{horizontal:'center',vertical:'center'}}; },
            b0:   function(){ return {fill:{patternType:'solid',fgColor:{rgb:'FFD7D7'}},alignment:{horizontal:'center',vertical:'center'}}; },
            b25:  function(){ return {fill:{patternType:'solid',fgColor:{rgb:'FFF2CC'}},alignment:{horizontal:'center',vertical:'center'}}; },
            b50:  function(){ return {fill:{patternType:'solid',fgColor:{rgb:'D9EAD3'}},alignment:{horizontal:'center',vertical:'center'}}; },
            b100: function(){ return {fill:{patternType:'solid',fgColor:{rgb:'93C47D'}},alignment:{horizontal:'center',vertical:'center'}}; }
        };

        function colLetter(n){var s='';while(n>0){var r2=(n-1)%26;s=String.fromCharCode(65+r2)+s;n=Math.floor((n-1)/26);}return s;}
        function coordR(r,c){return colLetter(max_r-r)+(max_c-c);}
        function dn(b){return b.nom_fr||b.nom||'';}
        function mc(v,s){return {v:(v===null||v===undefined)?'':v,t:typeof v==='number'?'n':'s',s:s};}

        var cI=originalPlaced.filter(function(b){return b.type==='Culturel';});
        var cO=optimized.filter(function(b){return b.type==='Culturel';});

        function cultRecv(prod,cults){
            var total=0;
            for(var i=0;i<cults.length;i++){
                var cu=cults[i],ray=cu.rayonnement||0;if(!ray)continue;
                var r0=cu.r,c0=cu.c,r1=r0+cu.rows-1,c1=c0+cu.cols-1,found=false;
                outer:for(var dr=0;dr<prod.rows;dr++)for(var dc=0;dc<prod.cols;dc++){
                    var rr=prod.r+dr,cc=prod.c+dc;
                    if(r0-ray<=rr&&rr<=r1+ray&&c0-ray<=cc&&cc<=c1+ray&&!(r0<=rr&&rr<=r1&&c0<=cc&&cc<=c1)){found=true;break outer;}
                }
                if(found)total+=cu.culture;
            }
            return total;
        }
        function boostLvl(cult,b){
            if(b.type!=='Producteur')return 0;
            if(b.boost100&&cult>=b.boost100)return 100;
            if(b.boost50&&cult>=b.boost50)return 50;
            if(b.boost25&&cult>=b.boost25)return 25;
            return 0;
        }

        function buildListe(placed,cults){
            var ws={};
            var hdrs=['Nom','Type','Niv.','Placé','Coord','Orient','Priorité','Placement',
                'Culture prod.','Rayon','Boost 25%','Boost 50%','Boost 100%',
                'Culture reçue','Boost atteint','Score boost'];
            hdrs.forEach(function(h,i){ws[colLetter(i+1)+'1']=mc(h,S.hdr());});
            var ri=2;
            placed.slice().sort(function(a,b2){var o={Culturel:0,Producteur:1,Neutre:2};return(o[a.type]||0)-(o[b2.type]||0)||dn(a).localeCompare(dn(b2));})
            .forEach(function(b){
                var cult=b.type==='Producteur'?cultRecv(b,cults):0;
                var bl=boostLvl(cult,b);
                var st=b.type==='Culturel'?S.cult():b.type==='Producteur'?S.prod():S.neut();
                var vals=[dn(b),b.type,b.level||1,'Oui',coordR(b.r,b.c),b.cols>=b.rows?'H':'V',
                    b.type==='Producteur'?b.priorite:'','Obligatoire',
                    b.type==='Culturel'?b.culture:'',b.type==='Culturel'?b.rayonnement:'',
                    b.type==='Producteur'?b.boost25:'',b.type==='Producteur'?b.boost50:'',b.type==='Producteur'?b.boost100:'',
                    Math.round(cult*10)/10,bl+'%',b.type==='Producteur'?Math.round(bl*b.priorite*100)/100:''];
                vals.forEach(function(v,i){ws[colLetter(i+1)+ri]=mc(v,st);});
                ri++;
            });
            var pc={};placed.forEach(function(b){pc[b.nom]=(pc[b.nom]||0)+1;});
            buildings_def.forEach(function(bdef){
                var miss=bdef.nombre-(pc[bdef.nom]||0);
                for(var mi=0;mi<miss;mi++){
                    var v2=[dn(bdef),bdef.type,'','Non','','','','Obligatoire','','','','','','','',''];
                    v2.forEach(function(v,i){ws[colLetter(i+1)+ri]=mc(v,S.miss());});
                    ri++;
                }
            });
            ws['!ref']='A1:'+colLetter(hdrs.length)+ri;
            ws['!cols']=[32,12,5,6,10,7,8,11,13,7,9,9,10,13,12,11].map(function(w){return{wch:w};});
            return ws;
        }

        XLSX.utils.book_append_sheet(wb,buildListe(originalPlaced,cI),'Liste batiments initiale');
        XLSX.utils.book_append_sheet(wb,buildListe(optimized,cO),'Liste batiments optimisee');

        // Synthèse
        var ws2={};
        var h2=['Type de batiment','Priorite','Avant 0%','Avant 25%','Avant 50%','Avant 100%','Apres 0%','Apres 25%','Apres 50%','Apres 100%'];
        h2.forEach(function(h,i){ws2[colLetter(i+1)+'1']=mc(h,S.hdr());});
        function bC(placed,cults){var r={};placed.forEach(function(b){if(b.type!=='Producteur')return;var nm=dn(b),c=cultRecv(b,cults),bl=boostLvl(c,b);if(!r[nm])r[nm]={0:0,25:0,50:0,100:0,prio:0};if(b.priorite>r[nm].prio)r[nm].prio=b.priorite;r[nm][bl]++;});return r;}
        var av=bC(originalPlaced,cI),ap=bC(optimized,cO);
        var allN=Object.keys(Object.assign({},av,ap)).sort();
        var ri2=2;
        allN.forEach(function(nm){
            var a=av[nm]||{0:0,25:0,50:0,100:0,prio:0},b2=ap[nm]||{0:0,25:0,50:0,100:0,prio:0};
            var sts=[null,null,S.b0(),S.b25(),S.b50(),S.b100(),S.b0(),S.b25(),S.b50(),S.b100()];
            [nm,a.prio||b2.prio,a[0],a[25],a[50],a[100],b2[0],b2[25],b2[50],b2[100]].forEach(function(v,i){ws2[colLetter(i+1)+ri2]=mc(v,sts[i]||{});});
            ri2++;
        });
        ri2+=2;
        var sAv=0,sAp=0;
        originalPlaced.forEach(function(b){if(b.type==='Producteur'&&b.priorite>0)sAv+=boostLvl(cultRecv(b,cI),b)*b.priorite;});
        optimized.forEach(function(b){if(b.type==='Producteur'&&b.priorite>0)sAp+=boostLvl(cultRecv(b,cO),b)*b.priorite;});
        var dl=Math.round((sAp-sAv)*100)/100;
        ws2['A'+ri2]=mc('Score global',{font:{bold:true}});
        ws2['D'+ri2]=mc(Math.round(sAv*100)/100,{font:{bold:true}});
        ws2['H'+ri2]=mc(Math.round(sAp*100)/100,{font:{bold:true}});
        ws2['I'+ri2]=mc(dl,{font:{bold:true,color:{rgb:dl>=0?'006400':'CC0000'}}});
        // Paramètres en bas de la Synthèse
        ri2+=3;
        var modeLabels={'priority':'Avec réduction','no_reduction':'Sans réduction','none':'Export seul'};
        ws2['A'+ri2]=mc('Paramètres d\'optimisation',S.hdr()); ws2['B'+ri2]=mc('',S.hdr());
        for(var ci_=2;ci_<=10;ci_++)ws2[colLetter(ci_)+ri2]=mc('',S.hdr());
        ri2++;
        ws2['A'+ri2]=mc('Mode',{font:{bold:true}});
        ws2['B'+ri2]=mc(modeLabels[optim_mode]||optim_mode||'',{});
        ri2++;
        var catNames={'Barracks':'Casernes','Farm':'Fermes','Home':'Maisons','Workshop':'Ateliers'};
        var protectedList=Object.keys(protected_cats||{}).filter(function(c){return protected_cats[c];}).map(function(c){return catNames[c]||c;});
        ws2['A'+ri2]=mc('Protéger',{font:{bold:true}});
        ws2['B'+ri2]=mc(protectedList.join(', ')||'Aucun',{});
        ri2+=2;
        ws2['A'+ri2]=mc('Bâtiment',S.hdr()); ws2['B'+ri2]=mc('Priorité',S.hdr());
        ri2++;
        var sortedPrios=Object.keys(prio_par_type||{}).sort(function(a,b){
            return (getLocaName(a)||a).localeCompare(getLocaName(b)||b);
        });
        sortedPrios.forEach(function(nom){
            var pv=prio_par_type[nom];
            if(pv===undefined||pv===null)return;
            var nomFR=getLocaName(nom)||cleanBuildingName(nom)||nom;
            var pvStr=pv==='max'?'max':String(pv);
            var pvStyle=pv==='max'?{font:{bold:true,color:{rgb:'CC6600'}}}:
                        pv>0?{font:{bold:true}}:{font:{color:{rgb:'888888'}}};
            ws2['A'+ri2]=mc(nomFR,{});
            ws2['B'+ri2]=mc(pvStr,pvStyle);
            ri2++;
        });
        ws2['!ref']='A1:'+colLetter(10)+(ri2+1);
        ws2['!cols']=[28,8,8,8,8,8,8,8,8,8].map(function(w){return{wch:w};});
        XLSX.utils.book_append_sheet(wb,ws2,'Synthese');

        // Déplacements
        var ws3={};
        ['#','Batiment','Position initiale','Position finale'].forEach(function(h,i){ws3[colLetter(i+1)+'1']=mc(h,S.hdr());});

        // Construire la liste des déplacements bruts
        var origPool={};originalPlaced.forEach(function(b){if(!origPool[b.nom])origPool[b.nom]=[];origPool[b.nom].push([b.r,b.c]);});
        var pool2={};Object.keys(origPool).forEach(function(n){pool2[n]=origPool[n].slice();});
        var rawMoves=[];
        optimized.forEach(function(b){
            var p=pool2[b.nom];if(!p||!p.length)return;
            var mi=-1;for(var i=0;i<p.length;i++){if(p[i][0]===b.r&&p[i][1]===b.c){mi=i;break;}}
            if(mi!==-1){p.splice(mi,1);return;}
            var bi2=0,bd=Infinity;for(var i=0;i<p.length;i++){var d=Math.abs(p[i][0]-b.r)+Math.abs(p[i][1]-b.c);if(d<bd){bd=d;bi2=i;}}
            var op=p.splice(bi2,1)[0];
            rawMoves.push({nom:dn(b),from:[op[0],op[1]],to:[b.r,b.c],rows:b.rows,cols:b.cols});
        });

        // Tri topologique : d'abord les déplacements dont la destination est libre
        // (pas occupée par une autre position initiale)
        function cellsOf(r,c,rows,cols){
            var cells=[];
            for(var dr=0;dr<rows;dr++)for(var dc=0;dc<cols;dc++)cells.push((r+dr)+'|'+(c+dc));
            return cells;
        }
        // Ensemble des cases initiales occupées
        var initOccupied={};
        rawMoves.forEach(function(mv){
            cellsOf(mv.from[0],mv.from[1],mv.rows,mv.cols).forEach(function(k){initOccupied[k]=true;});
        });

        var sorted=[], remaining=rawMoves.slice(), maxIter=rawMoves.length*2;
        while(remaining.length>0&&maxIter-->0){
            var progress=false;
            for(var ri_=0;ri_<remaining.length;ri_++){
                var mv=remaining[ri_];
                if(!mv.from||!mv.to){
                    // fromReserve : destination seulement, toujours plaçable
                    sorted.push(mv);
                    remaining.splice(ri_,1);
                    progress=true;
                    break;
                }
                // Vérifier si la destination est libre
                var destCells=cellsOf(mv.to[0],mv.to[1],mv.rows,mv.cols);
                var blocked=destCells.some(function(k){return initOccupied[k];});
                if(!blocked){
                    sorted.push(mv);
                    cellsOf(mv.from[0],mv.from[1],mv.rows,mv.cols).forEach(function(k){delete initOccupied[k];});
                    remaining.splice(ri_,1);
                    progress=true;
                    break;
                }
            }
            // Si aucun déplacement possible (cycle) : mettre en réserve
            if(!progress&&remaining.length>0){
                var mv0=remaining.shift();
                if(mv0.from&&mv0.to){
                    sorted.push({nom:mv0.nom, from:mv0.from, to:null, rows:mv0.rows, cols:mv0.cols, reserve:true});
                    cellsOf(mv0.from[0],mv0.from[1],mv0.rows,mv0.cols).forEach(function(k){delete initOccupied[k];});
                    remaining.push({nom:mv0.nom, from:null, to:mv0.to, rows:mv0.rows, cols:mv0.cols, fromReserve:true});
                } else {
                    sorted.push(mv0);
                }
            }
        }
        sorted=sorted.concat(remaining);

        var ri3=2;
        if(sorted.length===0){ws3['A2']=mc('Aucun deplacement.',{});}
        else sorted.forEach(function(mv,idx){
            var from_str = mv.fromReserve ? '⭐ Réserve'
                         : (mv.from ? coordR(mv.from[0],mv.from[1]) : '?');
            var to_str   = mv.reserve    ? '⭐ Réserve (temporaire)'
                         : (mv.to ? coordR(mv.to[0],mv.to[1]) : '?');
            var st = mv.reserve||mv.fromReserve ? S.b25() : S.move();
            [idx+1,mv.nom,from_str,to_str].forEach(function(v,i){ws3[colLetter(i+1)+ri3]=mc(v,st);});
            ri3++;
        });
        ws3['!ref']='A1:D'+Math.max(ri3,3);
        ws3['!cols']=[5,32,16,16].map(function(w){return{wch:w};});
        XLSX.utils.book_append_sheet(wb,ws3,'Deplacements');

        // Terrain
        function buildTerrain(placed,cults){
            var ws={},merges=[],used={};
            var grid={};
            placed.forEach(function(b){for(var dr=0;dr<b.rows;dr++)for(var dc=0;dc<b.cols;dc++)grid[(b.r+dr)+'|'+(b.c+dc)]=b;});
            var cw=[],rh=[];
            for(var ci2=0;ci2<max_r;ci2++)cw.push({wch:14});
            for(var ri4=0;ri4<max_c;ri4++)rh.push({hpt:20});
            for(var r=0;r<max_r;r++){
                for(var c=0;c<max_c;c++){
                    var tCol=max_r-r,tRow=max_c-c;
                    var ref=colLetter(tCol)+tRow;
                    if(terrain_grid[r][c]==='X'){
                        ws[ref]=mc('X',S.bx());
                    }else{
                        var b=grid[r+'|'+c];
                        if(!b)continue;
                        var st=b.type==='Culturel'?S.cult():b.type==='Producteur'?S.prod():S.neut();
                        if(b.r===r&&b.c===c){
                            var cult2=b.type==='Producteur'?cultRecv(b,cults):0;
                            var bl2=boostLvl(cult2,b);
                            var label=dn(b)+(b.type==='Producteur'&&bl2>0?' +'+bl2+'%':'');
                            var mcs=Math.max(1,tCol-b.rows+1),mrs=Math.max(1,tRow-b.cols+1);
                            var overlap=false;
                            for(var mr=mrs;mr<=tRow&&!overlap;mr++)for(var mc2=mcs;mc2<=tCol&&!overlap;mc2++)if(used[mr+'|'+mc2])overlap=true;
                            if(!overlap){
                                var lref=colLetter(mcs)+mrs;
                                var lst=b.type==='Culturel'?S.cult():b.type==='Producteur'?S.prod():S.neut();
                                if(bl2>0){lst=Object.assign({},lst);lst.font={bold:true};}
                                lst=Object.assign({},lst,{alignment:{horizontal:'center',vertical:'center',wrapText:true}});
                                ws[lref]=mc(label,lst);
                                if(b.rows>1||b.cols>1)merges.push({s:{r:mrs-1,c:mcs-1},e:{r:tRow-1,c:tCol-1}});
                                for(var mr2=mrs;mr2<=tRow;mr2++)for(var mc3=mcs;mc3<=tCol;mc3++)used[mr2+'|'+mc3]=true;
                            }else{ws[ref]=mc(label,st);}
                        }
                    }
                }
            }
            ws['!ref']='A1:'+colLetter(max_r)+max_c;
            ws['!merges']=merges;ws['!cols']=cw;ws['!rows']=rh;
            return ws;
        }
        XLSX.utils.book_append_sheet(wb,buildTerrain(optimized,cO),'Terrain optimise');
        XLSX.utils.book_append_sheet(wb,buildTerrain(originalPlaced,cI),'Terrain initial');
        return wb;
    }

    function downloadExcel(wb){
        var wbout=XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true});
        var blob=new Blob([wbout],{type:'application/octet-stream'});
        var url=URL.createObjectURL(blob);
        var a=document.createElement('a');a.href=url;
        var now=new Date(),pad=function(n){return n<10?'0'+n:''+n;};
        var ts=now.getFullYear()+''+pad(now.getMonth()+1)+''+pad(now.getDate())+'_'+pad(now.getHours())+''+pad(now.getMinutes());
        a.download='roc_ville_optimisee_'+ts+'.xlsx';
        document.body.appendChild(a);a.click();document.body.removeChild(a);
        URL.revokeObjectURL(url);
        log('\u2705 Excel t\u00e9l\u00e9charg\u00e9','#a8e6a3');
        updatePanel();
    }

    function getProducteurTypes(){
        var seen={}, result=[];
        allBuildings.forEach(function(b){
            if(seen[b.name])return; seen[b.name]=true;
            // Exclure les maisons de marins (nom FR contient "marin")
            var _nomFR=(getLocaName(b.name)||'').toLowerCase();
            if(_nomFR.indexOf('marin')!==-1)return;
            var info=getInfo(b.name);
            if(info.t100===0){var cat=csvBuildingCategory(b.name);if(cat!=='Farm'&&cat!=='Home'&&cat!=='Barracks')return;}
            var cat2=csvBuildingCategory(b.name); if(!cat2)return;
            result.push({nom:b.name,nomFR:getLocaName(b.name)||cleanBuildingName(b.name),cat:cat2});
        });
        var seenAff={},deduped=[];
        result.forEach(function(x){if(!seenAff[x.nomFR]){seenAff[x.nomFR]=true;deduped.push(x);}});
        var catOrder={Barracks:0,Farm:1,Home:2,Workshop:3};
        deduped.sort(function(a,b){return(catOrder[a.cat]||4)-(catOrder[b.cat]||4)||a.nomFR.localeCompare(b.nomFR);});
        return deduped;
    }

        function panelShowOptim(){
        var panel = document.getElementById('roc-export');
        if (!panel) return;

        // Test étape par étape
        panel.innerHTML = '<div style="color:#a8e6a3;padding:10px;font-size:12px;">Étape 1: getProducteurTypes...</div>';

        var prodTypes;
        try {
            prodTypes = getProducteurTypes();
        } catch(e) {
            panel.innerHTML = '<div style="color:#f08080;padding:10px;">ERR getProducteurTypes: '+e.message+'</div>';
            return;
        }

        panel.innerHTML = '<div style="color:#a8e6a3;padding:10px;font-size:12px;">Étape 2 OK: '+prodTypes.length+' types<br>Construction HTML...</div>';

        var catLabel = {Barracks:'🏰 Casernes', Farm:'🌾 Fermes', Home:'🏠 Maisons', Workshop:'⚒️ Ateliers'};
        var bycat = {};
        try {
            prodTypes.forEach(function(p){ if(!bycat[p.cat])bycat[p.cat]=[]; bycat[p.cat].push(p); });
        } catch(e) {
            panel.innerHTML = '<div style="color:#f08080;padding:10px;">ERR bycat: '+e.message+'</div>';
            return;
        }

        var prioRows = '';
        try {
            ['Barracks','Farm','Home','Workshop'].forEach(function(cat){
                if(!bycat[cat]) return;
                prioRows += '<div style="color:#a8e6a3;font-size:11px;font-weight:bold;margin:6px 0 2px;">'+catLabel[cat]+'</div>';
                prioRows += '<div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:4px;">';
                bycat[cat].forEach(function(p){
                    var cur = modalPrios[p.nom] !== undefined ? modalPrios[p.nom] : 0;
                    var opts = '';
                    for(var v=0;v<=10;v++) opts += '<option value="'+v+'"'+(cur===v?' selected':'')+'>'+v+'</option>';
                    opts += '<option value="max"'+(cur==='max'?' selected':'')+'>max</option>';
                    var nm = (p.nomFR||'').replace(/&/g,'&amp;').replace(/</g,'&lt;');
                    prioRows += '<div style="display:flex;flex-direction:column;align-items:center;background:#0d1f0d;padding:4px;border-radius:4px;min-width:78px;">'
                        +'<span style="font-size:9px;color:#ccc;text-align:center;margin-bottom:2px;">'+nm+'</span>'
                        +'<select id="prio_'+p.nom+'" style="background:#1a2a1a;color:#e0e0e0;border:1px solid #2a6a2a;border-radius:3px;padding:2px;font-size:11px;width:100%;">'
                        +opts+'</select></div>';
                });
                prioRows += '</div>';
            });
        } catch(e) {
            panel.innerHTML = '<div style="color:#f08080;padding:10px;">ERR prioRows: '+e.message+'</div>';
            return;
        }

        panel.innerHTML = '<div style="color:#a8e6a3;padding:10px;font-size:12px;">Étape 3 OK: HTML construit ('+prioRows.length+' chars)<br>Insertion...</div>';

        // Cases à cocher "protéger" par catégorie (pour mode sans réduction)
        var catKeys=['Barracks','Farm','Home','Workshop'];
        var catProtected=window._rocProtected||{Barracks:true,Farm:true,Home:true,Workshop:true};
        var protectHtml='<div id="roc-protect-zone"><div style="color:#a8e6a3;font-size:11px;font-weight:bold;margin:6px 0 2px;">Protéger (sans réduction)</div>';
        protectHtml+='<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px;">';
        catKeys.forEach(function(cat){
            if(!bycat[cat])return;
            var checked=catProtected[cat]?'checked':'';
            protectHtml+='<label style="display:flex;align-items:center;gap:3px;font-size:11px;cursor:pointer;">'
                +'<input type="checkbox" id="protect_'+cat+'" '+checked+'>'+catLabel[cat]+'</label>';
        });
        protectHtml+='</div></div>'; // ferme flex + roc-protect-zone

        var modeHtml = protectHtml;
        modeHtml += '<div style="color:#a8e6a3;font-size:11px;font-weight:bold;margin:4px 0 2px;">Mode</div>';
        ['priority','no_reduction','none'].forEach(function(v,i){
            var labels=['Avec réduction','Sans réduction','Export seul'];
            modeHtml += '<label style="display:flex;align-items:center;gap:4px;font-size:11px;margin-bottom:2px;cursor:pointer;">'
                +'<input type="radio" name="roc_mode_p" value="'+v+'"'+(modalMode===v?' checked':'')+'>'+labels[i]+'</label>';
        });

        try {
            panel.innerHTML =
                '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">'
                +'<strong style="color:#a8e6a3;font-size:12px;">⚙️ Optimiseur</strong>'
                +'<span id="roc-cancel-span" style="background:#333;color:#aaa;border-radius:3px;padding:2px 8px;font-size:13px;cursor:pointer;">✕</span></div>'
                +'<div style="font-size:9px;color:#888;margin-bottom:4px;">0=ignoré · 1-9=priorité · max=boost 100%</div>'
                + prioRows
                +'<div style="color:#a8e6a3;font-size:11px;font-weight:bold;margin:6px 0 2px;">Mode</div>'
                + modeHtml
                +'<div id="roc-launch-zone" style="margin-top:8px;">'
                +'<div style="display:flex;gap:6px;">'
                +'<span id="roc-cancel2-span" style="flex:1;background:#444;color:#fff;border-radius:6px;padding:10px;text-align:center;cursor:pointer;font-size:12px;">Annuler</span>'
                +'<span id="roc-lancer-span" style="flex:2;background:#1a6a2a;color:#fff;border-radius:6px;padding:10px;text-align:center;cursor:pointer;font-size:13px;font-weight:bold;">🚀 Lancer</span>'
                +'</div></div>';
        } catch(e) {
            panel.innerHTML = '<div style="color:#f08080;padding:10px;">ERR innerHTML: '+e.message+'</div>';
            return;
        }

        panel.style.maxHeight = '90vh';
        panel.style.overflowY = 'auto';

        // Listeners directs sur les boutons
        var _cancel = document.getElementById('roc-cancel-span');
        var _cancel2 = document.getElementById('roc-cancel2-span');
        var _lancer = document.getElementById('roc-lancer-span');

        function doCancel(e){ e.stopPropagation(); e.preventDefault(); updatePanel(); }
        function doLancer(e){ e.stopPropagation(); e.preventDefault(); startCountdown(); }

        if(_cancel){ _cancel.addEventListener('touchstart', doCancel, {passive:false}); _cancel.addEventListener('click', doCancel); }
        if(_cancel2){ _cancel2.addEventListener('touchstart', doCancel, {passive:false}); _cancel2.addEventListener('click', doCancel); }
        if(_lancer){ _lancer.addEventListener('touchstart', doLancer, {passive:false}); _lancer.addEventListener('click', doLancer); }


    }

    function startCountdown(){
        // Sauvegarder les priorités
        document.querySelectorAll('[id^="prio_"]').forEach(function(sel){
            var nom=sel.id.replace('prio_','');
            modalPrios[nom]=(sel.value==='max')?'max':(parseInt(sel.value)||0);
        });
        var modeEl=document.querySelector('input[name="roc_mode_p"]:checked');
        if(modeEl) modalMode=modeEl.value;
        // Lire les cases protéger
        window._rocProtected={};
        ['Barracks','Farm','Home','Workshop'].forEach(function(cat){
            var cb=document.getElementById('protect_'+cat);
            window._rocProtected[cat]=cb?cb.checked:true;
        });

        var lz=document.getElementById('roc-launch-zone');
        if(!lz) return;
        lz.innerHTML='<div style="text-align:center;background:#0a1f0a;border-radius:6px;padding:8px;">'
            +'<div style="color:#888;font-size:10px;">Lancement dans</div>'
            +'<div id="roc-cd" style="color:#a8e6a3;font-size:32px;font-weight:bold;">5</div>'
            +'<span id="roc-cancel-cd" style="background:#555;color:#fff;border-radius:4px;padding:4px 12px;font-size:11px;cursor:pointer;display:inline-block;">Annuler</span></div>';

        var _cc = document.getElementById('roc-cancel-cd');
        if(_cc){
            _cc.addEventListener('touchstart', function(e){e.stopPropagation();e.preventDefault();clearInterval(window._rocTimer);updatePanel();},{passive:false});
            _cc.addEventListener('click', function(e){e.stopPropagation();e.preventDefault();clearInterval(window._rocTimer);updatePanel();});
        }

        var n=5;
        window._rocTimer=setInterval(function(){
            n--;
            var el=document.getElementById('roc-cd');
            if(!el){clearInterval(window._rocTimer);return;}
            if(n<=0){clearInterval(window._rocTimer);el.textContent='🚀';setTimeout(function(){runOptimization();},300);}
            else el.textContent=n;
        },1000);
    }
    function runOptimization(){
        var panel = document.getElementById('roc-export');
        function show(msg){ if(panel) panel.innerHTML = '<div style="color:#a8e6a3;padding:10px;font-size:12px;line-height:1.5;">'+msg+'</div>'; }

        show('Lecture des priorités...');

        var prio_par_type={}, boost100_req=new Set(), aff_to_prio={};
        Object.keys(modalPrios).forEach(function(nom){
            var val=modalPrios[nom];
            var pv=val==='max'?10:(parseInt(val)||0);
            prio_par_type[nom]=pv;
            if(val==='max') boost100_req.add(nom);
            var nomAff=getLocaName(nom)||cleanBuildingName(nom);
            if(aff_to_prio[nomAff]===undefined||pv>aff_to_prio[nomAff]) aff_to_prio[nomAff]=pv;
            if(val==='max') aff_to_prio['__req__'+nomAff]=true;
        });
        allBuildings.forEach(function(b){
            if(prio_par_type[b.name]!==undefined) return;
            var nomAff=getLocaName(b.name)||cleanBuildingName(b.name);
            if(aff_to_prio[nomAff]!==undefined){
                prio_par_type[b.name]=aff_to_prio[nomAff];
                if(aff_to_prio['__req__'+nomAff]) boost100_req.add(b.name);
            }
        });

        var data;
        try { data = buildOptimData(prio_par_type, boost100_req); }
        catch(e){ show('Erreur terrain: '+e.message); return; }
        if(!data){ show('Terrain introuvable — visitez votre ville'); return; }

        show('Terrain OK ('+data.max_r+'×'+data.max_c+')<br>Placement des bâtiments...');
        setTimeout(function(){
            var placed, originalPlaced;
            try{
                placed = placeMissingBuildings(
                    data.placed, data.buildings_def,
                    data.terrain_grid, data.max_r, data.max_c);
                originalPlaced = data.placed.map(function(b){return Object.assign({},b);});
            }catch(e){ show('Erreur placement: '+e.message); return; }

            // Calculer boost_min pour les catégories protégées (tous modes sauf export seul)
            if(modalMode!=='none'){
                var cults0=placed.filter(function(b){return b.type==='Culturel';});
                var prot=window._rocProtected||{};
                var nProtected=0, nSkipped=0;
                placed.forEach(function(b){
                    if(b.type!=='Producteur')return;
                    var cat=csvBuildingCategory(b.nom||b.name||'');
                    if(!cat||!prot[cat]){nSkipped++;return;}
                    var c=cultureReceived(b,cults0);
                    b.boost_min=boostLevel(c,b);
                    nProtected++;
                });
                log('Protection: '+nProtected+' protégés, '+nSkipped+' ignorés. prot='+JSON.stringify(prot),'#f0c040');
                // Log les fermes rurales spécifiquement
                placed.forEach(function(b){
                    if(b.nom&&b.nom.toUpperCase().indexOf('RURAL')!==-1){
                        log('FermeRurale nom='+b.nom+' cat='+csvBuildingCategory(b.nom)+' boost_min='+b.boost_min,'#f0c040');
                    }
                });
            }

            if(modalMode==='none'){
                finishOptim(placed, originalPlaced, data, boost100_req, show, prio_par_type, modalMode, window._rocProtected||{});
                return;
            }

            show('Placement OK ('+placed.length+' bâtiments)<br>Optimisation en cours...');
            setTimeout(function(){
                var optimized;
                try{
                    optimized = optimizeMultiswap(
                        placed, data.terrain_grid, data.max_r, data.max_c,
                        function(f){ show('Optimisation: '+Math.round(f*100)+'%...'); },
                        60000, modalMode);
                }catch(e){ show('Erreur optimisation: '+e.message); return; }
                finishOptim(optimized, originalPlaced, data, boost100_req, show, prio_par_type, modalMode, window._rocProtected||{});
            }, 100);
        }, 100);
    }
    function finishOptim(optimized, originalPlaced, data, boost100_req, show, prio_par_type, optim_mode, protected_cats){
        show('Génération Excel...');
        setTimeout(function(){
            var wb;
            try{
                wb = buildExcelOutput(
                    optimized, originalPlaced,
                    data.terrain_grid, data.max_r, data.max_c,
                    data.buildings_def, boost100_req,
                    prio_par_type, optim_mode, protected_cats);
            }catch(e){
                show('Erreur Excel: '+e.message);
                setTimeout(function(){ updatePanel(); }, 5000);
                return;
            }
            show('✅ Terminé ! Téléchargement...');
            downloadExcel(wb);
        }, 100);
    }
    // ═══════════════════════════════════════════════════════════════
    // SECTION 13 : PANEL PRINCIPAL
    // ═══════════════════════════════════════════════════════════════

    function updatePanel(){var p=document.getElementById('roc-export');if(p)renderContent(p);}

    function renderContent(panel){
        var cityCount={};allBuildings.forEach(function(b){cityCount[b.city]=(cityCount[b.city]||0)+1;});

        var html='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">'
            +'<strong style="color:#a8e6a3;font-size:13px;">🏙️ RoC Optimiseur v6.4</strong>'
            +'<span data-roc="minimize" style="background:#333;color:#aaa;border-radius:3px;padding:3px 10px;font-size:14px;display:inline-block;text-decoration:none;">—</span>'
            +'</div>';

        var s1=catalogLoaded?'✓ '+Object.keys(CATALOG).length+' types':'⏳ catalogue';
        var s2=startupLoaded?'✓ startup':'⏳ startup';
        var s3=locaLoaded?'✓ '+Object.keys(LOCA).length+' trad.':'⏳ loca';
        html+='<div style="color:#888;font-size:10px;margin-bottom:6px;">'+s1+' | '+s2+' | '+s3+'</div>';

        if(allBuildings.length>0){
            html+='<div style="background:#0a1f0a;border:1px solid #2a4a2a;border-radius:6px;padding:8px;margin-bottom:8px;">'
                +'<div style="color:#a8e6a3;font-weight:bold;margin-bottom:4px;">✓ '+allBuildings.length+' bâtiments collectés</div>';
            Object.keys(cityCount).forEach(function(city){html+='<div style="color:#ccc;font-size:11px;">  📍 '+esc(city)+' : '+cityCount[city]+'</div>';});
            html+='</div>';
            html+='<span id="roc-optim-span" style="width:100%;background:#1a4a6a;color:#fff;border-radius:6px;padding:10px;font-size:12px;font-weight:bold;margin-bottom:6px;text-align:center;display:block;touch-action:manipulation;box-sizing:border-box;cursor:pointer;">⚙️ Optimiser &amp; Exporter Excel</span>';
        }else{
            html+='<div style="color:#888;font-size:11px;text-align:center;padding:10px;">En attente de données…<br>Visitez votre ville pour commencer.</div>';
        }

        
        

        if(logs.length>0){
            html+='<div style="background:#0a1628;border-radius:4px;padding:5px;max-height:80px;overflow-y:auto;">';
            logs.slice(-6).forEach(function(l){html+='<div style="color:'+l.color+';font-size:10px;">'+esc(l.msg)+'</div>';});
            html+='</div>';
        }
        panel.innerHTML=html;
        // Listener direct sur le bouton Optimiser
        var _os=document.getElementById('roc-optim-span');
        if(_os){
            _os.addEventListener('touchstart',function(e){
                e.stopPropagation();e.preventDefault();
                _os.style.background='#ff0000'; // feedback visuel
                setTimeout(function(){panelShowOptim();},100);
            },{passive:false});
            _os.addEventListener('click',function(e){
                e.stopPropagation();e.preventDefault();
                panelShowOptim();
            });
        }
    }

    function init(){
        var panel=document.createElement('div');
        panel.id='roc-export';
        panel.style.cssText='position:fixed;bottom:20px;right:20px;z-index:2147483647;background:#1a2a1a;color:#e0e0e0;font-family:monospace;font-size:12px;padding:12px;border-radius:10px;box-shadow:0 4px 24px rgba(0,0,0,0.8);width:310px;max-height:160px;overflow-y:auto;border:1px solid #2a6a2a;pointer-events:all;touch-action:manipulation;-webkit-tap-highlight-color:rgba(0,0,0,0);';
        document.body.appendChild(panel);
        // Propagation gérée par _rocDocHandler sur window
        renderContent(panel);

        var btn=document.createElement('button');
        btn.id='roc-export-btn';btn.textContent='🏙️';
        btn.style.cssText='display:none;position:fixed;bottom:20px;right:20px;z-index:999999;background:#1a2a1a;color:#fff;border:2px solid #2a6a2a;border-radius:50%;width:46px;height:46px;font-size:22px;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,0.8);';
        btn.setAttribute('onclick',"document.getElementById('roc-export').style.display='block';this.style.display='none';");
        document.body.appendChild(btn);

        if(!window.cityMuseumData)window.cityMuseumData={};

        // Delegation: un seul listener permanent sur le panel
        // Panel actions via window._roc (href javascript:)

        setTimeout(function(){loadCatalog(function(){processPending();updatePanel();});},100);
    }

    // Listener global unique pour tous les éléments data-roc
    // Fonctionne en capture sur document — avant le jeu
    function _rocHandleAction(actionName) {
        if (!actionName || !window._roc) return;
        var fn = window._roc[actionName];
        if (fn) fn();
    }
    (function() {
        function _rocDocHandler(e) {
            var t = e.target;
            // Remonter jusqu'à trouver un data-roc
            while (t && t !== document.body) {
                var action = t.getAttribute && t.getAttribute('data-roc');
                if (action) {
                    e.preventDefault();
                    e.stopImmediatePropagation();
                    e.stopPropagation();
                    _rocHandleAction(action);
                    return;
                }
                t = t.parentElement;
            }
        }
        // Écouter sur window ET document en capture — touchstart, mousedown, click
        window.addEventListener('touchstart',   _rocDocHandler, {capture:true, passive:false});
        window.addEventListener('mousedown',    _rocDocHandler, {capture:true});
        window.addEventListener('click',        _rocDocHandler, {capture:true});
    })();

    // Exposer les actions sur window IMMEDIATEMENT (avant init)
    // href="javascript:window._roc.xxx()" fonctionne sur iOS Safari/Tampermonkey
    window._roc = {
        run:       function() { runOptimization(); },
        do_optim:  function() { runOptimization(); },
        start_countdown: function() {
            // Lire les prios maintenant
            document.querySelectorAll('[id^="prio_"]').forEach(function(sel){
                var nom=sel.id.replace('prio_','');
                var val=sel.value;
                modalPrios[nom]=(val==='max')?'max':(parseInt(val)||0);
            });
            var modeEl=document.querySelector('input[name="roc_mode_p"]:checked');
            if(modeEl) modalMode=modeEl.value;
            // Afficher compte à rebours dans la zone launch
            var lz=document.getElementById('roc-launch-zone');
            if(lz) lz.innerHTML='<div style="text-align:center;background:#0a1f0a;border-radius:6px;padding:8px;">'
                +'<div style="color:#888;font-size:10px;">Lancement dans</div>'
                +'<div id="roc-cd" style="color:#a8e6a3;font-size:32px;font-weight:bold;">5</div>'
                +'<span data-roc="cancel_countdown" style="background:#555;color:#fff;border-radius:4px;'
                +'padding:4px 12px;font-size:11px;cursor:pointer;">Annuler</span></div>';
            var n=5;
            window._rocTimer=setInterval(function(){
                n--;
                var el=document.getElementById('roc-cd');
                if(!el){clearInterval(window._rocTimer);return;}
                if(n<=0){clearInterval(window._rocTimer);el.textContent='🚀';
                    setTimeout(function(){window._roc.do_optim();},300);}
                else el.textContent=n;
            },1000);
        },
        cancel_countdown: function() {
            if(window._rocTimer){clearInterval(window._rocTimer);window._rocTimer=null;}
            var lz=document.getElementById('roc-launch-zone');
            if(lz) lz.innerHTML='<div style="display:flex;gap:6px;">'
                +'<span data-roc="cancel_optim" style="flex:1;background:#444;color:#fff;border-radius:6px;'
                +'padding:10px;text-align:center;cursor:pointer;font-size:12px;">Annuler</span>'
                +'<span data-roc="start_countdown" style="flex:2;background:#1a6a2a;color:#fff;border-radius:6px;'
                +'padding:10px;text-align:center;cursor:pointer;font-size:13px;font-weight:bold;">🚀 Lancer</span>'
                +'</div>';
        },
        cancel_optim: function() {
            if (window._rocTimer) { clearInterval(window._rocTimer); window._rocTimer = null; }
            var ov = document.getElementById('roc-overlay');
            if (ov) ov.remove();
            updatePanel();
        },
        close:     function() { var w=document.getElementById('roc-iframe-wrap'); if(w) w.remove(); },
        cancel:    function() { var w=document.getElementById('roc-iframe-wrap'); if(w) w.remove(); },
        optim:     function() { panelShowOptim(); },
        cult:   function() {
            if(!window._rocWasmMem) findWasmMemory();
            if(!window._rocWasmMem){log('WASM non disponible','#f08080');return;}
            var mem32=new Uint32Array(window._rocWasmMem.buffer),NULL32=4294967295,cands=[];
            for(var i=3;i<mem32.length-2;i++){var v=mem32[i];if(v>=10&&v<=5000&&mem32[i-1]===NULL32&&mem32[i-2]===NULL32&&mem32[i+1]===NULL32&&mem32[i-3]===429)cands.push({offset:i*4,value:v});}
            log('Candidats: '+cands.length,'#ffaa00');
            if(cands.length===1){window._rocCultureOffset=cands[0].offset;window._rocCultureValue=cands[0].value;try{localStorage.setItem('_rocCultureOffset',cands[0].offset);}catch(err){}var pid=window._rocLastPlayerId||0;if(pid&&window.cityMuseumData&&window.cityMuseumData[pid])window.cityMuseumData[pid].culture=cands[0].value;log('Culture CityHall: '+cands[0].value,'#00ff88');updatePanel();}
        },

        minimize: function() {
            var p=document.getElementById('roc-export'), b=document.getElementById('roc-export-btn');
            if(p) p.style.display='none'; if(b) b.style.display='block';
        }
    };

    if(document.body)init();
    else document.addEventListener('DOMContentLoaded',init);

})();
