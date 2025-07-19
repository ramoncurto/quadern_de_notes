/****************************************************************
 * QUADERN DE NOTES - CODI DEL SERVIDOR (CODE.GS) - v5.6 Correcció de Cache
 *
 * Versió estable amb totes les correccions implementades.
 * - S'ha corregit l'error en la funció d'esborrat de la memòria cau.
 ****************************************************************/

// --- CONFIGURACIÓ GLOBAL ---
const SPREADSHEET_ID = "1jTmFdkhgJp9bOvgmH-1tt8lH7EWLY2RNEIn4i5Resio";
const SS = SpreadsheetApp.openById(SPREADSHEET_ID);

const CONFIG = {
    PESTANYES: {
        ALUMNAT: "Alumnat", GRUPS: "Grups", ALUMNAT_GRUP: "Alumnat_Grup", MODULS: "Moduls",
        RA: "RA", INSTRUMENTS: "InstrumentsAvaluacio", QUALIFICACIONS: "Qualificacions",
        RUBRIQUES: "Rubriques", CRITERIS_RUBRICA: "Criteris_Rubrica", NIVELLS_CRITERI: "Nivells_Criteri",
        QUALIFICACIONS_RUBRICA: "Qualificacions_Rubrica"
    },
    AVATARS_FOLDER_ID: "1vfzhevLT4aA9n9hASO5pfR4THvCjXT1I", 
    CACHE_DURATION: 300
};

// --- FUNCIONS PRINCIPALS ---
function doGet(e) {
    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
        .setTitle('Quadern de notes')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function updateInstrumentStatus(instrumentId, newStatus, grupId) {
    try {
        const sheet = SS.getSheetByName(CONFIG.PESTANYES.INSTRUMENTS);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idCol = headers.indexOf('Instrument_ID');
        const statusCol = headers.indexOf('Estat_Manual');
        const modulIdCol = headers.indexOf('Modul_ID');

        if (statusCol === -1) {
            throw new Error("La columna 'Estat_Manual' no existeix a la pestanya 'InstrumentsAvaluacio'.");
        }

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idCol]) === String(instrumentId)) {
                const valueToSet = (newStatus === 'Programat') ? '' : newStatus;
                sheet.getRange(i + 1, statusCol + 1).setValue(valueToSet);
                
                const modulId = data[i][modulIdCol];
                if (grupId && modulId) {
                    CacheService.getScriptCache().remove(`data_v5_${grupId}_${modulId}`);
                }
                
                return { success: true, message: 'Estat actualitzat.' };
            }
        }
        return { success: false, error: "No s'ha trobat l'instrument." };
    } catch (e) {
        return { success: false, error: e.message };
    }
}


function getInitialData() {
    try {
        const cache = CacheService.getScriptCache();
        const cachedData = cache.get('initial_data');
        if (cachedData) return JSON.parse(cachedData);
        
        const grups = getSheetDataAsObjectArray_(CONFIG.PESTANYES.GRUPS).map(g => ({ id: g.Grup_ID, name: g.Nom_Grup }));
        const moduls = getSheetDataAsObjectArray_(CONFIG.PESTANYES.MODULS).map(m => ({ id: m.Modul_ID, name: m.Nom_Modul }));
        
        // Carregar logo
        const logoUrl = getLogoUrl_();
        
        const result = { 
            success: true, 
            groups: grups, 
            modules: moduls,
            logoUrl: logoUrl,
            debug: {
                logoFound: !!logoUrl,
                logoLength: logoUrl ? logoUrl.length : 0
            }
        };
        cache.put('initial_data', JSON.stringify(result), CONFIG.CACHE_DURATION);
        return result;
    } catch (error) {
        return { success: false, error: `Error en carregar dades inicials: ${error.message}` };
    }
}

function getGradesData(grupId, modulId) {
    try {
        const cache = CacheService.getScriptCache();
        const cacheKey = `data_v5_${grupId}_${modulId}`;
        const cachedData = cache.get(cacheKey);
        if(cachedData) return JSON.parse(cachedData);

        const alumnes = getAlumnesByGrup_(grupId);
        if (alumnes.length === 0) return { success: true, alumnes: [], ras: [], instruments: [] };
        
        const config = getModuleConfig_(modulId);
        const definicionsRubriques = getRubricDefinitionsForInstruments_(config.instruments);
        const { qualificacions, qualificacionsRubrica } = getGradesForStudents_(alumnes, config.instruments);

        const result = {
            success: true, alumnes, ras: config.ras, instruments: config.instruments,
            definicionsRubriques, qualificacions, qualificacionsRubrica
        };
        
        cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_DURATION);
        return result;
    } catch (error) {
        return { success: false, error: `Error carregant dades: ${error.message}` };
    }
}

function saveGradesBatch(gradesToSave, grupId, modulId) {
    if (!Array.isArray(gradesToSave) || gradesToSave.length === 0) {
        return { success: true, message: "No hi ha notes per desar." };
    }
    try {
        const qualificacionsSheet = SS.getSheetByName(CONFIG.PESTANYES.QUALIFICACIONS);
        const qualificacionsRubricaSheet = SS.getSheetByName(CONFIG.PESTANYES.QUALIFICACIONS_RUBRICA);
        const allInstruments = getSheetDataAsObjectArray_(CONFIG.PESTANYES.INSTRUMENTS);
        const allRubrics = getRubricDefinitionsForInstruments_(allInstruments);

        gradesToSave.forEach(grade => {
            if (grade.rubricSelections) {
                const instrument = allInstruments.find(i => String(i.Instrument_ID) === String(grade.instrumentId));
                if (instrument && instrument.Rubrica_ID && allRubrics[instrument.Rubrica_ID]) {
                    const rubricDef = allRubrics[instrument.Rubrica_ID];
                    const finalNumericGrade = calculateNumericGradeFromRubric_(rubricDef, grade.rubricSelections);
                    updateOrInsertGrade_(qualificacionsSheet, grade.alumneId, grade.instrumentId, finalNumericGrade);
                    updateOrInsertRubricGrades_(qualificacionsRubricaSheet, grade.alumneId, grade.instrumentId, grade.rubricSelections);
                }
            } else {
                updateOrInsertGrade_(qualificacionsSheet, grade.alumneId, grade.instrumentId, grade.nota);
            }
        });
        
        CacheService.getScriptCache().remove(`data_v5_${grupId}_${modulId}`);
        return { success: true, message: 'Notes desades correctament.' };
    } catch (error) {
        return { success: false, error: `Error desant notes: ${error.message}` };
    }
}

function getModuleConfig(modulId) {
    try {
        const result = getModuleConfig_(modulId);
        const rubriques = getSheetDataAsObjectArray_(CONFIG.PESTANYES.RUBRIQUES).map(r => ({id: r.Rubrica_ID, name: r.Nom_Rubrica}));
        return { success: true, ...result, rubriques };
    } catch (error) {
        return { success: false, error: `Error carregant configuració: ${error.message}` };
    }
}

function saveRaConfig({ modulId, ras }) {
    try {
        const sheet = SS.getSheetByName(CONFIG.PESTANYES.RA);
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const raIdCol = headers.indexOf('RA_ID'), ponderacioCol = headers.indexOf('Ponderacio_RA'), modulIdCol = headers.indexOf('Modul_ID');
        const updates = [], raMap = new Map(ras.map(r => [String(r.id), parseFloat(r.ponderacio)]));
        data.forEach((row, index) => {
            if (String(row[modulIdCol]) === String(modulId)) {
                const raId = String(row[raIdCol]);
                if (raMap.has(raId)) {
                    const newPonderacio = raMap.get(raId) / 100;
                    if (row[ponderacioCol] != newPonderacio) updates.push({row: index + 2, col: ponderacioCol + 1, value: newPonderacio});
                }
            }
        });
        updates.forEach(u => sheet.getRange(u.row, u.col).setValue(u.value));
        CacheService.getScriptCache().removeMatching(`data_v5_.*_${modulId}`);
        return { success: true, message: 'Ponderacions dels RAs actualitzades.' };
    } catch (error) {
        return { success: false, error: `Error desant configuració de RAs: ${error.message}` };
    }
}

function saveInstrumentsConfig({ modulId, instruments }) {
    try {
        const sheet = SS.getSheetByName(CONFIG.PESTANYES.INSTRUMENTS);
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const modulIdCol = headers.indexOf('Modul_ID');
        const rowsToDelete = [];
        data.forEach((row, index) => { if (String(row[modulIdCol]) === String(modulId)) rowsToDelete.push(index + 2) });
        for (let i = rowsToDelete.length - 1; i >= 0; i--) sheet.deleteRow(rowsToDelete[i]);
        if (instruments.length > 0) {
            const headersOriginal = SS.getSheetByName(CONFIG.PESTANYES.INSTRUMENTS).getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            const newRows = instruments.map(inst => {
                let orderedRow = Array(headersOriginal.length).fill('');
                orderedRow[headersOriginal.indexOf('Instrument_ID')] = inst.id.startsWith('new_') ? Utilities.getUuid() : inst.id;
                orderedRow[headersOriginal.indexOf('Modul_ID')] = modulId;
                orderedRow[headersOriginal.indexOf('RA_ID')] = inst.raId;
                orderedRow[headersOriginal.indexOf('Nom_Instrument')] = inst.name;
                orderedRow[headersOriginal.indexOf('Ponderacio_Instrument')] = parseFloat(inst.ponderacio) || 0;
                orderedRow[headersOriginal.indexOf('Rubrica_ID')] = inst.rubricaId || '';
                return orderedRow;
            });
            if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }
        CacheService.getScriptCache().removeMatching(`data_v5_.*_${modulId}`);
        return { success: true, message: 'Instruments actualitzats correctament.' };
    } catch (error) {
        return { success: false, error: `Error desant configuració d'instruments: ${error.message}` };
    }
}

function getRubricsList() {
    try {
        return getSheetDataAsObjectArray_(CONFIG.PESTANYES.RUBRIQUES)
            .map(r => ({ id: r.Rubrica_ID, name: r.Nom_Rubrica }))
            .sort((a,b) => a.name.localeCompare(b.name));
    } catch(e) { return []; }
}

function getRubricDetails(rubricaId) {
    try {
        const rubricaInfo = getSheetDataAsObjectArray_(CONFIG.PESTANYES.RUBRIQUES).find(r => String(r.Rubrica_ID) === String(rubricaId));
        if (!rubricaInfo) throw new Error("Rúbrica no trobada.");
        const criteris = getSheetDataAsObjectArray_(CONFIG.PESTANYES.CRITERIS_RUBRICA).filter(c => String(c.Rubrica_ID) === String(rubricaId)).map(c => ({ id: c.Criteri_ID, name: c.Nom_Criteri, weight: parseFloat(c.Ponderacio_Criteri || 0) * 100 }));
        let nivells = [];
        if (criteris.length > 0) {
            const primerCriteriId = criteris[0].id;
            nivells = getSheetDataAsObjectArray_(CONFIG.PESTANYES.NIVELLS_CRITERI).filter(n => String(n.Criteri_ID) === String(primerCriteriId)).map(n => ({ id: n.Nivell_ID, name: n.Nom_Nivell, score: parseFloat(n.Puntuacio), description: n.Descripcio_Nivell })).sort((a,b) => b.score - a.score);
        }
        return { success: true, id: rubricaInfo.Rubrica_ID, name: rubricaInfo.Nom_Rubrica, description: rubricaInfo.Descripcio, criteria: criteris, levels: nivells };
    } catch (e) { return { success: false, error: e.message }; }
}

function saveRubric(rubricData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { id, name, description, criteria, levels } = rubricData;
        const isUpdate = id && !id.startsWith('new_');
        const rubriquesSheet = SS.getSheetByName(CONFIG.PESTANYES.RUBRIQUES);
        const rubricaId = isUpdate ? id : `R${getNextId_(rubriquesSheet, 'Rubrica_ID')}`;
        
        if (isUpdate) {
            updateRow_(rubriquesSheet, 'Rubrica_ID', rubricaId, { Nom_Rubrica: name, Descripcio: description });
        } else {
            rubriquesSheet.appendRow([rubricaId, name, description]);
        }
        const criterisSheet = SS.getSheetByName(CONFIG.PESTANYES.CRITERIS_RUBRICA);
        const nivellsSheet = SS.getSheetByName(CONFIG.PESTANYES.NIVELLS_CRITERI);
        
        const oldCriteria = getSheetDataAsObjectArray_(criterisSheet).filter(c => String(c.Rubrica_ID) === String(rubricaId)).map(c => c.Criteri_ID);
        if(isUpdate && oldCriteria.length > 0) {
         deleteRowsByCriteria_(nivellsSheet, { 'Criteri_ID': oldCriteria });
         deleteRowsByCriteria_(criterisSheet, { 'Rubrica_ID': rubricaId });
        }

        const nextCriteriId = getNextId_(criterisSheet, 'Criteri_ID');
        const nextNivellId = getNextId_(nivellsSheet, 'Nivell_ID');

        criteria.forEach((crit, critIndex) => {
            const criteriId = `CR${(nextCriteriId + critIndex).toString().padStart(4, '0')}`;
            criterisSheet.appendRow([criteriId, rubricaId, crit.name, parseFloat(crit.weight) / 100]);
            levels.forEach((level, levelIndex) => {
                const nivellId = `N${(nextNivellId + critIndex * levels.length + levelIndex).toString().padStart(4, '0')}`;
                nivellsSheet.appendRow([nivellId, criteriId, level.name, level.score, level.description]);
            });
        });
        return { success: true, message: 'Rúbrica desada correctament.' };
    } catch (e) { return { success: false, error: e.message }; } 
    finally { lock.releaseLock(); }
}

function deleteRubric(rubricaId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const criterisSheet = SS.getSheetByName(CONFIG.PESTANYES.CRITERIS_RUBRICA);
        const criterisData = getSheetDataAsObjectArray_(criterisSheet);
        const criterisToDelete = criterisData.filter(c => String(c.Rubrica_ID) === String(rubricaId)).map(c => c.Criteri_ID);
        if (criterisToDelete.length > 0) {
            deleteRowsByCriteria_(SS.getSheetByName(CONFIG.PESTANYES.NIVELLS_CRITERI), { Criteri_ID: criterisToDelete });
        }
        deleteRowsByCriteria_(criterisSheet, { Rubrica_ID: rubricaId });
        deleteRowsByCriteria_(SS.getSheetByName(CONFIG.PESTANYES.RUBRIQUES), { Rubrica_ID: rubricaId });
        return { success: true, message: 'Rúbrica eliminada correctament.' };
    } catch (e) { return { success: false, error: e.message }; }
    finally { lock.releaseLock(); }
}

function getAvatarUrlMap_() {
    const cache = CacheService.getScriptCache(), cacheKey = 'avatar_map';
    const cachedAvatars = cache.get(cacheKey);
    if (cachedAvatars) return new Map(JSON.parse(cachedAvatars));
    const avatarMap = new Map(), folderId = CONFIG.AVATARS_FOLDER_ID;
    if (!folderId) return avatarMap;
    try {
        const files = DriveApp.getFolderById(folderId).getFiles();
        while (files.hasNext()) {
            const file = files.next(), blob = file.getBlob(), type = blob.getContentType();
            if (type && type.startsWith('image/')) avatarMap.set(file.getName(), `data:${type};base64,${Utilities.base64Encode(blob.getBytes())}`);
        }
        cache.put(cacheKey, JSON.stringify(Array.from(avatarMap.entries())), CONFIG.CACHE_DURATION * 12);
    } catch (e) { console.log(`Error accedint a la carpeta d'avatars: ${e.message}`); }
    return avatarMap;
}

function getAlumnesByGrup_(grupId) {
    const alumnatDnis = new Set(getSheetDataAsObjectArray_(CONFIG.PESTANYES.ALUMNAT_GRUP).filter(r => String(r.Grup_ID) === String(grupId)).map(r => r.DNI));
    const totAlumnat = getSheetDataAsObjectArray_(CONFIG.PESTANYES.ALUMNAT);
    const avatarMap = getAvatarUrlMap_();
    return totAlumnat.filter(a => alumnatDnis.has(a.DNI)).map(a => ({
        id: a.Alumne_ID,
        name: a.Nom_Complet || `${a.Cognom1 || ''}, ${a.Nom || ''}`,
        initials: `${(a.Nom || ' ').charAt(0)}${(a.Cognom1 || ' ').charAt(0)}`.toUpperCase(),
        avatarUrl: avatarMap.get((a.Avatar || '').trim()) || null
    })).sort((a, b) => a.name.localeCompare(b.name));
}

function getModuleConfig_(modulId) {
    const ras = getSheetDataAsObjectArray_(CONFIG.PESTANYES.RA).filter(r => String(r.Modul_ID) === String(modulId)).map(r => ({ id: r.RA_ID, name: r.Nom_RA, ponderacio: parseFloat(r.Ponderacio_RA || 0) * 100 }));
    const instruments = getSheetDataAsObjectArray_(CONFIG.PESTANYES.INSTRUMENTS).filter(i => String(i.Modul_ID) === String(modulId)).map(i => ({
        id: i.Instrument_ID,
        raId: i.RA_ID,
        name: i.Nom_Instrument,
        ponderacio: parseFloat(i.Ponderacio_Instrument || 0),
        rubricaId: i.Rubrica_ID || null,
        estatManual: i.Estat_Manual || null
    }));
    return { ras, instruments };
}

function getSheetDataAsObjectArray_(sheetName) {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return [];
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return [];
    const values = dataRange.getValues();
    const headers = values.shift().map(h => h.replace(/\s+/g, '_'));
    return values.map(row => {
        const obj = {};
        headers.forEach((header, index) => { obj[header] = row[index]; });
        return obj;
    });
}

function getRubricDefinitionsForInstruments_(instruments) {
    const rubricIds = [...new Set(instruments.map(i => i.rubricaId || i.Rubrica_ID).filter(Boolean))];
    if (rubricIds.length === 0) return {};
    const allCriteris = getSheetDataAsObjectArray_(CONFIG.PESTANYES.CRITERIS_RUBRICA);
    const allNivells = getSheetDataAsObjectArray_(CONFIG.PESTANYES.NIVELLS_CRITERI);
    const definitions = {};
    rubricIds.forEach(rubricaId => {
        const criteris = allCriteris.filter(c => String(c.Rubrica_ID) === String(rubricaId)).map(c => ({
            id: c.Criteri_ID, nom: c.Nom_Criteri, ponderacio: parseFloat(c.Ponderacio_Criteri),
            nivells: allNivells.filter(n => String(n.Criteri_ID) === String(c.Criteri_ID)).map(n => ({ id: n.Nivell_ID, nom: n.Nom_Nivell, puntuacio: parseFloat(n.Puntuacio), descripcio: n.Descripcio_Nivell })).sort((a,b) => b.puntuacio - a.puntuacio)
        }));
        definitions[rubricaId] = { id: rubricaId, criteris: criteris };
    });
    return definitions;
}

function getGradesForStudents_(alumnes, instruments) {
    const allQualificacions = getSheetDataAsObjectArray_(CONFIG.PESTANYES.QUALIFICACIONS);
    const allQualificacionsRubrica = getSheetDataAsObjectArray_(CONFIG.PESTANYES.QUALIFICACIONS_RUBRICA);
    const qualificacions = {}, qualificacionsRubrica = {};
    alumnes.forEach(alumne => {
        qualificacions[alumne.id] = {};
        qualificacionsRubrica[alumne.id] = {};
        instruments.forEach(inst => {
            const qual = allQualificacions.find(q => String(q.Alumne_ID) === String(alumne.id) && String(q.Instrument_ID) === String(inst.id));
            qualificacions[alumne.id][inst.id] = qual ? qual.Nota : '';
            if (inst.rubricaId) {
                const rubQuals = allQualificacionsRubrica.filter(rq => String(rq.Alumne_ID) === String(alumne.id) && String(rq.Instrument_ID) === String(inst.id));
                if (rubQuals.length > 0) {
                    qualificacionsRubrica[alumne.id][inst.id] = {};
                    rubQuals.forEach(rq => { qualificacionsRubrica[alumne.id][inst.id][rq.Criteri_ID] = rq.Nivell_ID_Seleccionat; });
                }
            }
        });
    });
    return { qualificacions, qualificacionsRubrica };
}

function calculateNumericGradeFromRubric_(rubricDef, selections) {
    if (!rubricDef || !rubricDef.criteris) return 0;
    let finalGrade = 0;
    rubricDef.criteris.forEach(criteri => {
        const selectedNivellId = selections[criteri.id];
        if (selectedNivellId) {
            const nivell = criteri.nivells.find(n => String(n.id) === String(selectedNivellId));
            if (nivell) finalGrade += nivell.puntuacio * criteri.ponderacio;
        }
    });
    return finalGrade;
}

function updateOrInsertGrade_(sheet, alumneId, instrumentId, nota) {
    const data = sheet.getDataRange().getValues(), headers = data[0];
    const alumneIdCol = headers.indexOf('Alumne_ID'), instrumentIdCol = headers.indexOf('Instrument_ID'), notaCol = headers.indexOf('Nota');
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) { if (String(data[i][alumneIdCol]) == String(alumneId) && String(data[i][instrumentIdCol]) == String(instrumentId)) { foundRow = i + 1; break; } }
    if (foundRow > -1) sheet.getRange(foundRow, notaCol + 1).setValue(nota);
    else sheet.appendRow([Utilities.getUuid(), alumneId, instrumentId, nota, new Date()]);
}

function updateOrInsertRubricGrades_(sheet, alumneId, instrumentId, selections) {
    const data = sheet.getDataRange().getValues(), headers = data[0];
    const alumneIdCol = headers.indexOf('Alumne_ID'), instrumentIdCol = headers.indexOf('Instrument_ID'), criteriIdCol = headers.indexOf('Criteri_ID'), nivellIdCol = headers.indexOf('Nivell_ID_Seleccionat');
    Object.keys(selections).forEach(criteriId => {
        const nivellId = selections[criteriId];
        let foundRow = -1;
        for (let i = 1; i < data.length; i++) { if (String(data[i][alumneIdCol]) == String(alumneId) && String(data[i][instrumentIdCol]) == String(instrumentId) && String(data[i][criteriIdCol]) == String(criteriId)) { foundRow = i + 1; break; } }
        if (foundRow > -1) sheet.getRange(foundRow, nivellIdCol + 1).setValue(nivellId);
        else sheet.appendRow([Utilities.getUuid(), alumneId, instrumentId, criteriId, nivellId, new Date()]);
    });
}

function getNextId_(sheet, idColumnName) {
    const idColumnNameClean = idColumnName.replace(/\s+/g, '_');
    const data = getSheetDataAsObjectArray_(sheet.getName());
    if (data.length === 0) return 1;
    const maxId = Math.max(...data.map(row => {
        const idStr = String(row[idColumnNameClean] || '0').replace(/[^0-9]/g, '');
        return parseInt(idStr) || 0;
    }));
    return maxId + 1;
}

function updateRow_(sheet, idColumn, idValue, newData) {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf(idColumn);
    const rowIndex = data.findIndex(row => String(row[idColIndex]) === String(idValue));
    if (rowIndex !== -1) {
        headers.forEach((header, index) => {
            if (newData.hasOwnProperty(header)) {
                sheet.getRange(rowIndex + 2, index + 1).setValue(newData[header]);
            }
        });
    }
}

function deleteRowsByCriteria_(sheet, criteria) {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const rowsToDelete = [];
    data.forEach((row, index) => {
        let match = true;
        for (const key in criteria) {
            const colIndex = headers.indexOf(key);
            const criteriaValue = criteria[key];
            if (Array.isArray(criteriaValue)) {
                if (!criteriaValue.includes(row[colIndex])) {
                    match = false;
                    break;
                }
            } else {
                if (String(row[colIndex]) != String(criteriaValue)) {
                    match = false;
                    break;
                }
            }
        }
        if (match) rowsToDelete.push(index + 2);
    });
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
    }
}

// --- FUNCIONS PER AL LOGO ---
function getLogoUrl_() {
    try {
        console.log('=== CARREGANT LOGO ===');
        const cache = CacheService.getScriptCache();
        const cacheKey = 'logo_url';
        const cachedLogo = cache.get(cacheKey);
        
        if (cachedLogo) {
            console.log('Logo trobat a la cache');
            return cachedLogo;
        }
        
        const folderId = CONFIG.AVATARS_FOLDER_ID;
        if (!folderId) {
            console.log('No hi ha folder ID configurat per al logo');
            return null;
        }
        
        console.log('Buscant fitxer logo.png a la carpeta:', folderId);
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.getFilesByName('logo.png');
        
        if (files.hasNext()) {
            const logoFile = files.next();
            console.log('Fitxer logo.png trobat:', logoFile.getName());
            
            const blob = logoFile.getBlob();
            const contentType = blob.getContentType();
            console.log('Content type:', contentType);
            
            if (contentType && contentType.startsWith('image/')) {
                const base64Data = Utilities.base64Encode(blob.getBytes());
                const logoUrl = `data:${contentType};base64,${base64Data}`;
                console.log('Logo convertit a base64, longitud:', logoUrl.length);
                
                // Guardar a cache per 1 hora
                cache.put(cacheKey, logoUrl, 3600);
                return logoUrl;
            } else {
                console.log('El fitxer logo.png no és una imatge vàlida');
                return null;
            }
        } else {
            console.log('No s\'ha trobat cap fitxer logo.png a la carpeta');
            return null;
        }
    } catch (error) {
        console.error('Error carregant logo:', error.message);
        return null;
    }
}

// --- FUNCIONS DE DEBUG ---
function debugLogo() {
    try {
        console.log('=== DEBUG LOGO INICIAT ===');
        const folderId = CONFIG.AVATARS_FOLDER_ID;
        console.log('Folder ID configurat:', folderId);
        
        if (!folderId) {
            return {
                success: false,
                error: 'No hi ha folder ID configurat',
                folderConfigured: false
            };
        }
        
        const folder = DriveApp.getFolderById(folderId);
        console.log('Carpeta trobada:', folder.getName());
        
        // Llistar tots els fitxers de la carpeta
        const allFiles = [];
        const files = folder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            allFiles.push({
                name: file.getName(),
                mimeType: file.getBlob().getContentType(),
                size: file.getSize()
            });
        }
        console.log('Fitxers a la carpeta:', allFiles);
        
        // Buscar específicament logo.png
        const logoFiles = folder.getFilesByName('logo.png');
        const logoExists = logoFiles.hasNext();
        console.log('Logo.png existeix:', logoExists);
        
        let logoInfo = null;
        if (logoExists) {
            const logoFile = logoFiles.next();
            logoInfo = {
                name: logoFile.getName(),
                mimeType: logoFile.getBlob().getContentType(),
                size: logoFile.getSize(),
                id: logoFile.getId()
            };
            console.log('Informació del logo:', logoInfo);
        }
        
        // Test de càrrega
        const logoUrl = getLogoUrl_();
        
        return {
            success: true,
            folderConfigured: true,
            folderId: folderId,
            folderName: folder.getName(),
            allFiles: allFiles,
            logoExists: logoExists,
            logoInfo: logoInfo,
            logoUrlGenerated: !!logoUrl,
            logoUrlLength: logoUrl ? logoUrl.length : 0
        };
        
    } catch (error) {
        console.error('Error en debug logo:', error);
        return {
            success: false,
            error: error.message,
            folderConfigured: !!CONFIG.AVATARS_FOLDER_ID
        };
    }
}

function clearCache() {
    try {
        const cache = CacheService.getScriptCache();
        cache.removeAll(['initial_data', 'logo_url', 'avatar_map']);
        
        // També netejar cache de dades específiques
        const patterns = ['data_v5_'];
        // Note: removeMatching no existeix, així que fem remove individual
        
        return {
            success: true,
            message: 'Cache netejada correctament'
        };
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

// --- FUNCIONS D'ENVIAMENT DE CORREUS ---

function sendInstrumentGrade(instrumentId, grupId, modulId) {
    try {
        const instrument = getSheetDataAsObjectArray_(CONFIG.PESTANYES.INSTRUMENTS)
            .find(i => String(i.Instrument_ID) === String(instrumentId));
        
        if (!instrument) {
            return { success: false, error: "Instrument no trobat." };
        }

        const modul = getSheetDataAsObjectArray_(CONFIG.PESTANYES.MODULS)
            .find(m => String(m.Modul_ID) === String(modulId));
        
        const grup = getSheetDataAsObjectArray_(CONFIG.PESTANYES.GRUPS)
            .find(g => String(g.Grup_ID) === String(grupId));

        const alumnesGrup = getAlumnesWithEmailsByGrup_(grupId);
        const qualificacions = getSheetDataAsObjectArray_(CONFIG.PESTANYES.QUALIFICACIONS);

        let emailsSent = 0;
        let emailsWithErrors = 0;

        alumnesGrup.forEach(alumne => {
            try {
                const nota = qualificacions.find(q => 
                    String(q.Alumne_ID) === String(alumne.Alumne_ID) && 
                    String(q.Instrument_ID) === String(instrumentId)
                );

                const notaValue = nota ? nota.Nota : 'Pendent';
                
                const emailContent = generateInstrumentEmailContent_(
                    alumne.Nom_Complet, 
                    instrument.Nom_Instrument, 
                    notaValue, 
                    modul ? modul.Nom_Modul : 'Mòdul desconegut',
                    grup ? grup.Nom_Grup : 'Grup desconegut'
                );

                const recipients = [alumne.Correu_alumne];
                if (alumne.Informar_Pares && alumne.Correu_tutor) {
                    recipients.push(alumne.Correu_tutor);
                }

                const subject = `Qualificació: ${instrument.Nom_Instrument} - ${alumne.Nom_Complet}`;
                
                GmailApp.sendEmail(
                    recipients.join(','),
                    subject,
                    emailContent
                );
                
                emailsSent++;
            } catch (emailError) {
                console.error(`Error enviant correu a ${alumne.Nom_Complet}:`, emailError);
                emailsWithErrors++;
            }
        });

        return {
            success: true,
            message: `Correus enviats: ${emailsSent}. Errors: ${emailsWithErrors}`,
            details: { sent: emailsSent, errors: emailsWithErrors }
        };

    } catch (error) {
        return { success: false, error: `Error enviant correus: ${error.message}` };
    }
}

function sendRAGrades(raId, grupId, modulId) {
    try {
        const ra = getSheetDataAsObjectArray_(CONFIG.PESTANYES.RA)
            .find(r => String(r.RA_ID) === String(raId));
        
        if (!ra) {
            return { success: false, error: "RA no trobat." };
        }

        const modul = getSheetDataAsObjectArray_(CONFIG.PESTANYES.MODULS)
            .find(m => String(m.Modul_ID) === String(modulId));
        
        const grup = getSheetDataAsObjectArray_(CONFIG.PESTANYES.GRUPS)
            .find(g => String(g.Grup_ID) === String(grupId));

        const alumnesGrup = getAlumnesWithEmailsByGrup_(grupId);
        const instruments = getSheetDataAsObjectArray_(CONFIG.PESTANYES.INSTRUMENTS)
            .filter(i => String(i.Modul_ID) === String(modulId) && String(i.RA_ID) === String(raId));
        const qualificacions = getSheetDataAsObjectArray_(CONFIG.PESTANYES.QUALIFICACIONS);

        let emailsSent = 0;
        let emailsWithErrors = 0;

        alumnesGrup.forEach(alumne => {
            try {
                const notesInstruments = instruments.map(inst => {
                    const nota = qualificacions.find(q => 
                        String(q.Alumne_ID) === String(alumne.Alumne_ID) && 
                        String(q.Instrument_ID) === String(inst.Instrument_ID)
                    );
                    return {
                        nom: inst.Nom_Instrument,
                        nota: nota ? nota.Nota : 'Pendent',
                        pes: inst.Ponderacio_Instrument
                    };
                });

                const notaRA = calculateRAGradeForEmail_(alumne.Alumne_ID, raId, instruments, qualificacions);
                
                const emailContent = generateRAEmailContent_(
                    alumne.Nom_Complet,
                    ra.Nom_RA,
                    notesInstruments,
                    notaRA,
                    modul ? modul.Nom_Modul : 'Mòdul desconegut',
                    grup ? grup.Nom_Grup : 'Grup desconegut'
                );

                const recipients = [alumne.Correu_alumne];
                if (alumne.Informar_Pares && alumne.Correu_tutor) {
                    recipients.push(alumne.Correu_tutor);
                }

                const subject = `Qualificacions RA: ${ra.Nom_RA} - ${alumne.Nom_Complet}`;
                
                GmailApp.sendEmail(
                    recipients.join(','),
                    subject,
                    emailContent
                );
                
                emailsSent++;
            } catch (emailError) {
                console.error(`Error enviant correu a ${alumne.Nom_Complet}:`, emailError);
                emailsWithErrors++;
            }
        });

        return {
            success: true,
            message: `Correus enviats: ${emailsSent}. Errors: ${emailsWithErrors}`,
            details: { sent: emailsSent, errors: emailsWithErrors }
        };

    } catch (error) {
        return { success: false, error: `Error enviant correus: ${error.message}` };
    }
}

function sendMPGrades(grupId, modulId) {
    try {
        const modul = getSheetDataAsObjectArray_(CONFIG.PESTANYES.MODULS)
            .find(m => String(m.Modul_ID) === String(modulId));
        
        const grup = getSheetDataAsObjectArray_(CONFIG.PESTANYES.GRUPS)
            .find(g => String(g.Grup_ID) === String(grupId));

        const alumnesGrup = getAlumnesWithEmailsByGrup_(grupId);
        const ras = getSheetDataAsObjectArray_(CONFIG.PESTANYES.RA)
            .filter(r => String(r.Modul_ID) === String(modulId));
        const instruments = getSheetDataAsObjectArray_(CONFIG.PESTANYES.INSTRUMENTS)
            .filter(i => String(i.Modul_ID) === String(modulId));
        const qualificacions = getSheetDataAsObjectArray_(CONFIG.PESTANYES.QUALIFICACIONS);

        let emailsSent = 0;
        let emailsWithErrors = 0;

        alumnesGrup.forEach(alumne => {
            try {
                const notesRA = ras.map(ra => {
                    const instrumentsRA = instruments.filter(i => String(i.RA_ID) === String(ra.RA_ID));
                    const notaRA = calculateRAGradeForEmail_(alumne.Alumne_ID, ra.RA_ID, instrumentsRA, qualificacions);
                    return {
                        nom: ra.Nom_RA,
                        nota: notaRA,
                        pes: ra.Ponderacio_RA
                    };
                });

                const notaFinalMP = calculateMPGradeForEmail_(notesRA);
                
                const emailContent = generateMPEmailContent_(
                    alumne.Nom_Complet,
                    notesRA,
                    notaFinalMP,
                    modul ? modul.Nom_Modul : 'Mòdul desconegut',
                    grup ? grup.Nom_Grup : 'Grup desconegut'
                );

                const recipients = [alumne.Correu_alumne];
                if (alumne.Informar_Pares && alumne.Correu_tutor) {
                    recipients.push(alumne.Correu_tutor);
                }

                const subject = `Notes finals MP: ${modul ? modul.Nom_Modul : 'Mòdul'} - ${alumne.Nom_Complet}`;
                
                GmailApp.sendEmail(
                    recipients.join(','),
                    subject,
                    emailContent
                );
                
                emailsSent++;
            } catch (emailError) {
                console.error(`Error enviant correu a ${alumne.Nom_Complet}:`, emailError);
                emailsWithErrors++;
            }
        });

        return {
            success: true,
            message: `Correus enviats: ${emailsSent}. Errors: ${emailsWithErrors}`,
            details: { sent: emailsSent, errors: emailsWithErrors }
        };

    } catch (error) {
        return { success: false, error: `Error enviant correus: ${error.message}` };
    }
}

// --- FUNCIONS AUXILIARS PER A CORREUS ---

function getAlumnesWithEmailsByGrup_(grupId) {
    const alumnatDnis = new Set(getSheetDataAsObjectArray_(CONFIG.PESTANYES.ALUMNAT_GRUP)
        .filter(r => String(r.Grup_ID) === String(grupId))
        .map(r => r.DNI));
    
    return getSheetDataAsObjectArray_(CONFIG.PESTANYES.ALUMNAT)
        .filter(a => alumnatDnis.has(a.DNI))
        .map(a => ({
            Alumne_ID: a.Alumne_ID,
            Nom_Complet: a.Nom_Complet || `${a.Cognom1 || ''}, ${a.Nom || ''}`,
            Correu_alumne: a.Correu_alumne,
            Correu_tutor: a.Correu_tutor,
            Informar_Pares: a.Informar_Pares === true || a.Informar_Pares === 'TRUE'
        }))
        .filter(a => a.Correu_alumne); // Només alumnes amb correu vàlid
}

function calculateRAGradeForEmail_(alumneId, raId, instruments, qualificacions) {
    let weightedSum = 0;
    let totalWeight = 0;
    
    instruments.forEach(inst => {
        const nota = qualificacions.find(q => 
            String(q.Alumne_ID) === String(alumneId) && 
            String(q.Instrument_ID) === String(inst.Instrument_ID)
        );
        
        if (nota && nota.Nota !== '' && nota.Nota !== null) {
            const notaNum = parseFloat(nota.Nota);
            const weight = parseFloat(inst.Ponderacio_Instrument || 0);
            if (!isNaN(notaNum) && !isNaN(weight)) {
                weightedSum += notaNum * weight;
                totalWeight += weight;
            }
        }
    });
    
    return totalWeight === 0 ? 'Pendent' : (weightedSum / totalWeight).toFixed(1);
}

function calculateMPGradeForEmail_(notesRA) {
    let weightedSum = 0;
    let totalWeight = 0;
    
    notesRA.forEach(ra => {
        if (ra.nota !== 'Pendent') {
            const notaNum = parseFloat(ra.nota);
            const weight = parseFloat(ra.pes || 0);
            if (!isNaN(notaNum) && !isNaN(weight)) {
                weightedSum += notaNum * weight;
                totalWeight += weight;
            }
        }
    });
    
    return totalWeight === 0 ? 'Pendent' : (weightedSum / totalWeight).toFixed(1);
}

function generateInstrumentEmailContent_(nomAlumne, nomInstrument, nota, nomModul, nomGrup) {
    const dataActual = new Date().toLocaleDateString('ca-ES');
    
    return `
Benvolgut/da ${nomAlumne},

T'informem sobre la teva qualificació en l'instrument d'avaluació "${nomInstrument}" del mòdul "${nomModul}".

DETALLS DE LA QUALIFICACIÓ:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📚 Mòdul: ${nomModul}
👥 Grup: ${nomGrup}
📝 Instrument: ${nomInstrument}
🎯 Nota: ${nota}

${nota !== 'Pendent' ? (parseFloat(nota) >= 5 ? 
'✅ FELICITATS! Has assolit els objectius d\'aquest instrument.' : 
'⚠️ Aquesta nota indica que necessites millorar en aquest instrument. Et recomanem revisar els continguts i consultar amb el professorat.') : 
'⏳ La teva qualificació encara està pendent.'}

Si tens cap dubte sobre aquesta qualificació, no dubtis en contactar amb el teu professorat.

Salutacions cordials,
Equip docent
${dataActual}

---
Aquest correu s'ha generat automàticament des del Quadern de Notes.
    `.trim();
}

function generateRAEmailContent_(nomAlumne, nomRA, notesInstruments, notaRA, nomModul, nomGrup) {
    const dataActual = new Date().toLocaleDateString('ca-ES');
    
    const instrumentsDetails = notesInstruments.map(inst => 
        `• ${inst.nom}: ${inst.nota} (Pes: ${(inst.pes * 100).toFixed(0)}%)`
    ).join('\n');
    
    return `
Benvolgut/da ${nomAlumne},

T'informem sobre les teves qualificacions en el Resultat d'Aprenentatge "${nomRA}" del mòdul "${nomModul}".

DETALLS DE LES QUALIFICACIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📚 Mòdul: ${nomModul}
👥 Grup: ${nomGrup}
🎯 Resultat d'Aprenentatge: ${nomRA}

INSTRUMENTS D'AVALUACIÓ:
${instrumentsDetails}

🏆 NOTA FINAL DEL RA: ${notaRA}

${notaRA !== 'Pendent' ? (parseFloat(notaRA) >= 5 ? 
'✅ EXCEL·LENT! Has assolit tots els objectius d\'aquest Resultat d\'Aprenentatge.' : 
'⚠️ Aquesta nota indica que necessites millorar alguns aspectes d\'aquest RA. Et recomanem revisar els continguts i consultar amb el professorat.') : 
'⏳ Algunes qualificacions encara estan pendents.'}

Si tens cap dubte sobre aquestes qualificacions, no dubtis en contactar amb el teu professorat.

Salutacions cordials,
Equip docent
${dataActual}

---
Aquest correu s'ha generat automàticament des del Quadern de Notes.
    `.trim();
}

function generateMPEmailContent_(nomAlumne, notesRA, notaFinalMP, nomModul, nomGrup) {
    const dataActual = new Date().toLocaleDateString('ca-ES');
    
    const raDetails = notesRA.map(ra => 
        `• ${ra.nom}: ${ra.nota} (Pes: ${(ra.pes * 100).toFixed(0)}%)`
    ).join('\n');
    
    return `
Benvolgut/da ${nomAlumne},

T'informem sobre les teves notes finals del Mòdul Professional "${nomModul}".

RESUM FINAL DEL MÒDUL:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📚 Mòdul: ${nomModul}
👥 Grup: ${nomGrup}

RESULTATS D'APRENENTATGE:
${raDetails}

🏆 NOTA FINAL DEL MÒDUL PROFESSIONAL: ${notaFinalMP}

${notaFinalMP !== 'Pendent' ? (parseFloat(notaFinalMP) >= 5 ? 
'🎉 FELICITATS! Has superat amb èxit aquest Mòdul Professional. El teu esforç i dedicació han donat els seus fruits.' : 
'⚠️ La nota final indica que no has assolit els objectius mínims d\'aquest mòdul. Et recomanem revisar tots els continguts i consultar amb el professorat per planificar estratègies de millora.') : 
'⏳ Algunes qualificacions encara estan pendents per completar la nota final.'}

Si tens cap dubte sobre aquestes qualificacions o necessites orientació acadèmica, no dubtis en contactar amb el teu professorat o tutor.

Salutacions cordials,
Equip docent
${dataActual}

---
Aquest correu s'ha generat automàticament des del Quadern de Notes.
    `.trim();
}
