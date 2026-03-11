// --- Operation Control ---
let isOperationCancelled = false;

function setupCancelButton() {
    const cancelButton = document.getElementById('cancelButton');
    if (cancelButton) {
        cancelButton.addEventListener('click', () => {
            isOperationCancelled = true;
            document.getElementById('progressText').textContent = 'Operation cancelled';
            document.getElementById('loading').style.display = 'none';
            document.getElementById('error').textContent = 'Operation was cancelled by user';
            document.getElementById('error').style.display = 'block';
        });
    }
}

// Initialize cancel button
document.addEventListener('DOMContentLoaded', setupCancelButton);

// --- Universal Fetch Function ---
async function fetchFromAPI(endpoint, fioriId, releaseId) {
    if (isOperationCancelled) {
        throw new Error('Operation cancelled by user');
    }
    const base = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')`;
    const response = await fetch(`${base}/${endpoint}?$format=json`);
    if (!response.ok) throw new Error(`HTTP error! Status: ${response.status}`);
    const data = await response.json();
    return data.d?.results || [];
}

// --- Wrappers for specific endpoints ---
// CHANGE 1: Each fetch wrapper is now called conditionally based on selected fields

async function fetchTechnicalCatalogs(fioriId, releaseId)    { return fetchFromAPI('SplitTechnicalCatalogs', fioriId, releaseId); }
async function fetchSemanticObjects(fioriId, releaseId)      { return fetchFromAPI('SplitAdditionalIntents', fioriId, releaseId); }
async function fetchRelatedApps(fioriId, releaseId)          { return fetchFromAPI('Related_Apps', fioriId, releaseId); }
async function fetchSpaces(fioriId, releaseId)               { return fetchFromAPI('SplitSpace', fioriId, releaseId); }
async function fetchPages(fioriId, releaseId)                { return fetchFromAPI('SplitPage', fioriId, releaseId); }
async function fetchTechnicalNames(fioriId, releaseId)       { return fetchFromAPI('ODataServices', fioriId, releaseId); }
async function fetchSemanticActions(fioriId, releaseId)      {
    const data = await fetchFromAPI('SplitAdditionalIntents', fioriId, releaseId);
    return data.map(item => ({
        SemanticObject: item.SemanticObject,
        SemanticAction: item.SemanticAction
    }));
}
async function fetchBusinessRoleNames(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(fioriId='${fioriId}',releaseId='${releaseId}',inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN')/SplitBusinessRole?$format=json`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP error! Status: ${response.status}`);
    const data = await response.json();
    return data.d?.results || [];
}
async function fetchBSPNames(fioriId, releaseId) {
    const data = await fetchFromAPI('ICFNodes', fioriId, releaseId);
    return data.map(item => ({
        ...item,
        BSPName: item.isAdditional === 1 ? `${item.BSPName}*` : item.BSPName
    }));
}
async function fetchAppDetails(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(fioriId='${fioriId}',releaseId='${releaseId}',inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN')?$format=json`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP error! Status: ${response.status}`);
    const data = await response.json();
    return data.d;
}

// --- Check if a field is selected ---
function isFieldSelected(fieldName) {
    const checkboxes = document.querySelectorAll('.excel-field');
    for (const cb of checkboxes) {
        if (cb.value === fieldName && cb.checked) return true;
    }
    return false;
}

// --- CHANGE 2: Check if an app has target mappings (semantic objects/actions) ---
async function hasTargetMappings(fioriId, releaseId) {
    try {
        const data = await fetchFromAPI('SplitAdditionalIntents', fioriId, releaseId);
        return data && data.length > 0;
    } catch (e) {
        return false;
    }
}

// --- Retry Utility ---
async function retryFetch(fetchFn, maxRetries = 1) {
    for (let i = 0; i < maxRetries; i++) {
        if (isOperationCancelled) {
            throw new Error('Operation cancelled by user');
        }
        try { return await fetchFn(); }
        catch (error) {
            if (i === maxRetries - 1) throw error;
            await new Promise(resolve => setTimeout(resolve, Math.pow(2, i) * 1000));
        }
    }
}

// --- Error Message ---
function showError(message) {
    const errorElement = document.getElementById('error');
    errorElement.textContent = message;
    errorElement.style.display = 'block';
    setTimeout(() => { errorElement.style.display = 'none'; }, 5000);
}

// --- Clipboard Copy ---
async function copyToClipboard(text) {
    try { await navigator.clipboard.writeText(text); return true; }
    catch { return false; }
}

// --- Data List Display Helper ---
function displayDataList(containerId, textContainerId, copyButtonId, successId, data, nameFn, detailsFn) {
    const container = document.getElementById(containerId);
    const textEl = document.getElementById(textContainerId);
    if (!container || !textEl) return;
    container.innerHTML = '';

    if (!data?.length) {
        container.innerHTML = `<p>No ${containerId} found</p>`;
        textEl.textContent = '';
        return;
    }

    const names = [...new Set(data.map(nameFn))];
    textEl.textContent = names.join('\n');
    textEl.style.display = 'block';

    const copyBtn = document.getElementById(copyButtonId);
    if (copyBtn) {
        copyBtn.onclick = async () => {
            if (await copyToClipboard(names.join('\n'))) {
                const success = document.getElementById(successId);
                if (success) {
                    success.style.display = 'inline';
                    setTimeout(() => (success.style.display = 'none'), 2000);
                }
            }
        };
    }

    data.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerHTML = detailsFn(item);
        container.appendChild(div);
    });
}

// --- Display Config ---
const displayConfig = {
    technicalCatalogs: {
        name: i => i.TechincalCatalog,
        details: i =>
            `<p><strong>${i.TechincalCatalog}</strong></p>
            ${i.TechincalCatalogDescription ? `<p class="data-item-details">Description: ${i.TechincalCatalogDescription}</p>` : ''}
            ${i.SystemAlias ? `<p class="data-item-details">System Alias: ${i.SystemAlias}</p>` : ''}`
    },
    semanticObjects: {
        name: i => i.SemanticObject,
        details: i =>
            `<p><strong>${i.SemanticObject}</strong></p>
            ${i.SemanticAction ? `<p class="data-item-details">Action: ${i.SemanticAction}</p>` : ''}
            ${i.MappingSignatureKeyVal ? `<p class="data-item-details">Parameters: ${i.MappingSignatureKeyVal}</p>` : ''}`
    },
    relatedApps: {
        name: i => i.FioriId,
        details: i =>
            `<p><strong>${i.FioriId}</strong></p>
            ${i.AppName ? `<p class="data-item-details">App Name: ${i.AppName}</p>` : ''}
            ${i.relationType ? `<p class="data-item-details">Relation Type: ${i.relationType}</p>` : ''}`
    },
    spaces: {
        name: i => i.SpaceName,
        details: i =>
            `<p><strong>${i.SpaceName}</strong></p>
            ${i.SpaceTitle ? `<p class="data-item-details">Title: ${i.SpaceTitle}</p>` : ''}
            ${i.SpaceDescription ? `<p class="data-item-details">Description: ${i.SpaceDescription}</p>` : ''}`
    },
    pages: {
        name: i => i.PageName,
        details: i =>
            `<p><strong>${i.PageName}</strong></p>
            ${i.PageTitle ? `<p class="data-item-details">Title: ${i.PageTitle}</p>` : ''}
            ${i.PageDescription ? `<p class="data-item-details">Description: ${i.PageDescription}</p>` : ''}`
    },
    technicalNames: {
        name: i => i.TechnicalName,
        details: i =>
            `<p><strong>${i.TechnicalName}</strong></p>
            ${i.NameSpace ? `<p class="data-item-details">Namespace: ${i.NameSpace}</p>` : ''}
            ${i.SoftwareComponentName ? `<p class="data-item-details">Software Component: ${i.SoftwareComponentName}</p>` : ''}`
    },
    bspNames: {
        name: b => b.BSPName,
        details: b =>
            `<p><strong>${b.BSPName}</strong>${b.isAdditional === 0 ? ' (Main BSP)' : ''}</p>
            ${b.AppName ? `<p class="data-item-details">App Name: ${b.AppName}</p>` : ''}
            ${b.BSPApplicationURL ? `<p class="data-item-details">URL: ${b.BSPApplicationURL}</p>` : ''}
            ${b.SAPUI5ComponentId ? `<p class="data-item-details">UI5 Component ID: ${b.SAPUI5ComponentId}</p>` : ''}`
    },
    businessRoles: {
        name: r => r.BusinessRoleName,
        details: r =>
            `<p><strong>${r.BusinessRoleName}</strong>${r.isLeading === "X" ? ' (Leading Role)' : ''}</p>
            ${r.BusinessRoleDescription ? `<p class="data-item-details">Description: ${r.BusinessRoleDescription}</p>` : ''}
            ${r.RoleID ? `<p class="data-item-details">Role ID: ${r.RoleID}</p>` : ''}`
    }
};

Object.keys(displayConfig).forEach(key => {
    window[`display${key.charAt(0).toUpperCase() + key.slice(1)}`] = data => {
        const cfg = displayConfig[key];
        displayDataList(key, `${key}Text`, `copy${key.charAt(0).toUpperCase() + key.slice(1)}`,
            `copy${key.charAt(0).toUpperCase() + key.slice(1)}Success`, data, cfg.name, cfg.details);
    };
});

// --- Processing Results Display ---
function updateProcessingResults(results) {
    const container = document.getElementById('processingResults');
    container.innerHTML = '';
    results.forEach(result => {
        const div = document.createElement('div');
        div.className = 'data-item';
        let statusClass = result.status === 'Success' ? 'success' : 'error';
        let statusText = result.status === 'Success' ?
            (result.isDeprecated ? 'Success (Deprecated)' : 'Success') : 'Error';
        let originIndicator = result.isOriginalRequest ? ' 🎯' : '';
        let html = `<p><strong>${result.fioriId}${originIndicator}</strong> - <span class="${statusClass}">${statusText}</span></p>`;
        if (result.status === 'Error') {
            html += `<p class="data-item-details">Error: ${result.error}</p>`;
        } else if (result.isDeprecated) {
            html += `<p class="data-item-details">⚠️ This application is deprecated in the selected release</p>`;
        }
        if (result.removedAsNoTargetMapping) {
            html += `<p class="data-item-details">🚫 Removed from related apps: no target mappings</p>`;
        }
        if (!result.isOriginalRequest) {
            html += `<p class="data-item-details">📎 Discovered as related app</p>`;
        }
        div.innerHTML = html;
        container.appendChild(div);
    });
}

// --- App Info Display ---
function displayAppInfo(appData) {
    const container = document.getElementById('appInfo');
    const appInfoSection = document.getElementById('appInfoSection');
    if (!appData || !container) return;
    let html = `
        <div class="app-title">${appData.Title || appData.AppName || 'Unknown App'}</div>
        <div class="app-meta">
            <div class="app-meta-item"><span class="app-meta-label">Fiori ID:</span> <span>${appData.fioriId || 'N/A'}</span></div>
            <div class="app-meta-item"><span class="app-meta-label">Release:</span> <span>${appData.ReleaseName || appData.releaseId || 'N/A'}</span></div>
            <div class="app-meta-item"><span class="app-meta-label">Application Type:</span> <span>${appData.ApplicationType || 'N/A'}</span></div>
            <div class="app-meta-item"><span class="app-meta-label">UI Technology:</span> <span>${appData.UITechnology || 'N/A'}</span></div>
        </div>`;
    if (appData.isPublished === 'Deprecated') {
        html += `<div class="deprecated-warning"><span class="deprecated-icon">⚠️</span>
            This application is marked as DEPRECATED in this release</div>`;
    }
    if (appData.ApplicationComponent) {
        html += `<div class="app-meta-item"><span class="app-meta-label">Application Component:</span>
            <span>${appData.ApplicationComponent} (${appData.ApplicationComponentText || ''})</span></div>`;
    }
    if (appData.BSPName) {
        html += `<div class="app-meta-item"><span class="app-meta-label">BSP Application:</span>
            <span>${appData.BSPName}</span></div>`;
    }
    if (appData.SAPUI5ComponentId) {
        html += `<div class="app-meta-item"><span class="app-meta-label">SAPUI5 Component ID:</span>
            <span>${appData.SAPUI5ComponentId}</span></div>`;
    }
    container.innerHTML = html;
    if (appInfoSection) appInfoSection.style.display = 'block';
}

// --- Get Selected Fields ---
function getSelectedFields() {
    return Array.from(document.querySelectorAll('.excel-field'))
        .filter(cb => cb.checked)
        .map(cb => cb.value);
}

// --- Enter Key Handler ---
function handleEnterKey(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const releaseId = document.getElementById('releaseId').value.trim();
        const fioriIds = document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0);
        const selectedFields = getSelectedFields();
        if (fioriIds.length === 0 || !releaseId) { showError('Please enter both Fiori App IDs and Release ID'); return; }
        if (selectedFields.length === 0) { showError('Please select at least one field to include in the Excel'); return; }
        document.getElementById('customDownloadButton').click();
    }
}

// --- Input Event Listener Setup ---
['fioriIds', 'releaseId'].forEach(id => {
    const el = document.getElementById(id);
    el.addEventListener('keypress', handleEnterKey);
    el.addEventListener('input', e => e.target.value = e.target.value.toUpperCase());
});

// =========================================================
// CHANGE 3: Sort results by Technical Catalog name
// Original apps sorted by their first catalog, related apps follow their parent
// =========================================================
function sortResultsByCatalog(results, originalFioriIds) {
    // Separate original and related
    const originalResults = results.filter(r => originalFioriIds.has(r.fioriId) && r.status === 'Success' && !r.isDeprecated);
    const relatedResults = results.filter(r => !originalFioriIds.has(r.fioriId) && r.status === 'Success' && !r.isDeprecated);
    const errorResults = results.filter(r => r.status === 'Error');
    const deprecatedResults = results.filter(r => r.status === 'Success' && r.isDeprecated);

    // Sort original apps by first technical catalog name
    originalResults.sort((a, b) => {
        const catA = (a.technicalCatalogs && a.technicalCatalogs.length > 0)
            ? (a.technicalCatalogs[0].TechincalCatalog || '') : '';
        const catB = (b.technicalCatalogs && b.technicalCatalogs.length > 0)
            ? (b.technicalCatalogs[0].TechincalCatalog || '') : '';
        return catA.localeCompare(catB);
    });

    // Sort related apps by first technical catalog name
    relatedResults.sort((a, b) => {
        const catA = (a.technicalCatalogs && a.technicalCatalogs.length > 0)
            ? (a.technicalCatalogs[0].TechincalCatalog || '') : '';
        const catB = (b.technicalCatalogs && b.technicalCatalogs.length > 0)
            ? (b.technicalCatalogs[0].TechincalCatalog || '') : '';
        return catA.localeCompare(catB);
    });

    // Combine: original first, then related, then deprecated, then errors
    return [...originalResults, ...relatedResults, ...deprecatedResults, ...errorResults];
}

// =========================================================
// CHANGE 1: Selective fetch based on checked fields
// =========================================================
async function fetchAppDataSelective(fioriId, releaseId, selectedFields) {
    // Always fetch app details (lightweight, needed for basic info)
    const appDetails = await fetchAppDetails(fioriId, releaseId);

    // Always fetch related apps (needed for graph traversal regardless of field selection)
    const relatedApps = await retryFetch(() => fetchRelatedApps(fioriId, releaseId));

    // Conditionally fetch based on selected fields
    const technicalNames = selectedFields.includes('OData Services')
        ? await retryFetch(() => fetchTechnicalNames(fioriId, releaseId)) : [];

    const businessRoles = selectedFields.includes('Business Roles')
        ? await retryFetch(() => fetchBusinessRoleNames(fioriId, releaseId)) : [];

    // BSP Name is needed for both 'BSP Name' field AND for the app details display
    const bspNames = selectedFields.includes('BSP Name')
        ? await retryFetch(() => fetchBSPNames(fioriId, releaseId)) : [];

    const technicalCatalogs = selectedFields.includes('Technical Catalogs')
        ? await retryFetch(() => fetchTechnicalCatalogs(fioriId, releaseId)) : [];

    const spaces = selectedFields.includes('Spaces')
        ? await retryFetch(() => fetchSpaces(fioriId, releaseId)) : [];

    const pages = selectedFields.includes('Pages')
        ? await retryFetch(() => fetchPages(fioriId, releaseId)) : [];

    // Semantic Objects and Semantic Actions share the same endpoint (SplitAdditionalIntents)
    // Fetch only once if either is selected, or if we need target mapping check
    let semanticObjects = [];
    let semanticActions = [];
    const needSemanticData = selectedFields.includes('Semantic Objects') || selectedFields.includes('Semantic Actions');
    if (needSemanticData) {
        const rawSemanticData = await retryFetch(() => fetchFromAPI('SplitAdditionalIntents', fioriId, releaseId));
        if (selectedFields.includes('Semantic Objects')) {
            semanticObjects = rawSemanticData;
        }
        if (selectedFields.includes('Semantic Actions')) {
            semanticActions = rawSemanticData.map(item => ({
                SemanticObject: item.SemanticObject,
                SemanticAction: item.SemanticAction
            }));
        }
    }

    return { appDetails, relatedApps, technicalNames, businessRoles, bspNames, technicalCatalogs, spaces, pages, semanticObjects, semanticActions };
}

// =========================================================
// Excel Generation (shared between main and deprecated downloads)
// =========================================================
function buildExcelWorkbook(validResults, selectedFields, originalFioriIds) {
    const wb = XLSX.utils.book_new();

    if (validResults.length === 0) return null;

    // Gather unique sets
    const uniqueSets = {
        businessRoles: new Set(), odataServices: new Set(), technicalCatalogs: new Set(),
        bspNames: new Set(), spaces: new Set(), pages: new Set(),
        relatedApps: new Set(), semanticObjects: new Set(), semanticActions: new Set()
    };

    validResults.forEach(result => {
        result.businessRoles.forEach(role => uniqueSets.businessRoles.add(role.BusinessRoleName));
        result.technicalNames.forEach(service => uniqueSets.odataServices.add(service.TechnicalName));
        result.technicalCatalogs.forEach(catalog => uniqueSets.technicalCatalogs.add(catalog.TechincalCatalog));
        if (result.appDetails.BSPName) uniqueSets.bspNames.add(result.appDetails.BSPName);
        result.bspNames.forEach(bsp => { if (bsp.BSPName) uniqueSets.bspNames.add(bsp.BSPName); });
        result.spaces.forEach(space => uniqueSets.spaces.add(space.SpaceName));
        result.pages.forEach(page => uniqueSets.pages.add(page.PageName));
        result.relatedApps.forEach(app => uniqueSets.relatedApps.add(app.FioriId));
        result.semanticObjects.forEach(obj => {
            uniqueSets.semanticObjects.add(obj.SemanticObject);
        });
        if (result.semanticActions) {
            result.semanticActions.forEach(sa => {
                if (sa.SemanticObject && sa.SemanticAction)
                    uniqueSets.semanticActions.add(`${sa.SemanticObject}:${sa.SemanticAction}`);
            });
        }
    });

    // Add App Type field if not present
    const allFields = selectedFields.includes('App Type') ? selectedFields : [...selectedFields, 'App Type'];

    const excelData = validResults.map(result => {
        const row = {};
        allFields.forEach(field => {
            if (field === 'App Type') {
                row[field] = originalFioriIds.has(result.fioriId) ? '✓' : '✗';
            } else {
                switch (field) {
                    case 'Fiori ID': row[field] = result.fioriId; break;
                    case 'App Title': row[field] = result.appDetails.Title || result.appDetails.AppName; break;
                    case 'Application Type': row[field] = result.appDetails.ApplicationType; break;
                    case 'UI Technology': row[field] = result.appDetails.UITechnology; break;
                    case 'Application Component': row[field] = result.appDetails.ApplicationComponent; break;
                    case 'BSP Name': {
                        let bspNamesList = [];
                        if (result.appDetails.BSPName) bspNamesList.push(result.appDetails.BSPName);
                        result.bspNames.forEach(bsp => {
                            if (bsp.BSPName && !bspNamesList.includes(bsp.BSPName)) bspNamesList.push(bsp.BSPName);
                        });
                        row[field] = bspNamesList.join('\n');
                        break;
                    }
                    case 'UI5 Component ID': row[field] = result.appDetails.SAPUI5ComponentId; break;
                    case 'Business Roles': row[field] = result.businessRoles.map(r => r.BusinessRoleName).join('\n'); break;
                    case 'OData Services': row[field] = result.technicalNames.map(s => s.TechnicalName).join('\n'); break;
                    case 'Technical Catalogs': row[field] = result.technicalCatalogs.map(c => c.TechincalCatalog).join('\n'); break;
                    case 'Spaces': row[field] = result.spaces.map(s => s.SpaceName).join('\n'); break;
                    case 'Pages': row[field] = result.pages.map(p => p.PageName).join('\n'); break;
                    case 'Related Apps': row[field] = result.relatedApps.map(a => a.FioriId).join('\n'); break;
                    case 'Semantic Objects': row[field] = result.semanticObjects.map(o => o.SemanticObject).join('\n'); break;
                    case 'Semantic Actions': row[field] = result.semanticActions ? result.semanticActions.map(sa => `${sa.SemanticObject}:${sa.SemanticAction}`).join('\n') : ''; break;
                    default: row[field] = '';
                }
            }
        });
        return row;
    });

    // Consolidated row
    const consolidatedRow = {};
    allFields.forEach(field => {
        switch (field) {
            case 'Fiori ID': consolidatedRow[field] = 'CONSOLIDATED'; break;
            case 'App Type': consolidatedRow[field] = '-'; break;
            case 'BSP Name': consolidatedRow[field] = Array.from(uniqueSets.bspNames).sort().join('\n'); break;
            case 'Business Roles': consolidatedRow[field] = Array.from(uniqueSets.businessRoles).sort().join('\n'); break;
            case 'OData Services': consolidatedRow[field] = Array.from(uniqueSets.odataServices).sort().join('\n'); break;
            case 'Technical Catalogs': consolidatedRow[field] = Array.from(uniqueSets.technicalCatalogs).sort().join('\n'); break;
            case 'Spaces': consolidatedRow[field] = Array.from(uniqueSets.spaces).sort().join('\n'); break;
            case 'Pages': consolidatedRow[field] = Array.from(uniqueSets.pages).sort().join('\n'); break;
            case 'Related Apps': consolidatedRow[field] = Array.from(uniqueSets.relatedApps).sort().join('\n'); break;
            case 'Semantic Objects': consolidatedRow[field] = Array.from(uniqueSets.semanticObjects).sort().join('\n'); break;
            case 'Semantic Actions': consolidatedRow[field] = Array.from(uniqueSets.semanticActions).sort().join('\n'); break;
            default: consolidatedRow[field] = '';
        }
    });
    excelData.push(consolidatedRow);

    const ws = XLSX.utils.json_to_sheet(excelData);
    const colWidthsMap = {
        'Fiori ID': 15, 'App Title': 40, 'Application Type': 20, 'UI Technology': 20,
        'Application Component': 30, 'BSP Name': 20, 'UI5 Component ID': 30,
        'Business Roles': 40, 'OData Services': 40, 'Technical Catalogs': 40,
        'App Type': 10, 'Spaces': 40, 'Pages': 40, 'Related Apps': 40,
        'Semantic Objects': 40, 'Semantic Actions': 40
    };
    ws['!cols'] = allFields.map(field => ({ wch: colWidthsMap[field] || 20 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Fiori Apps Data');
    return wb;
}

// --- Main Excel Generation (active/non-deprecated apps) ---
function generateExcel(results, selectedFields, originalFioriIds) {
    if (isOperationCancelled) return;

    const validResults = results.filter(r => r.status === 'Success' && !r.isDeprecated);
    if (validResults.length === 0) {
        showError('No valid (non-deprecated) apps found to generate Excel');
        return;
    }

    try {
        const wb = buildExcelWorkbook(validResults, selectedFields, originalFioriIds);
        if (!isOperationCancelled && wb) {
            const releaseId = document.getElementById('releaseId').value.trim();
            XLSX.writeFile(wb, `Fiori_Apps_Data_${releaseId}.xlsx`);
        }
    } catch (error) {
        console.error('Error generating Excel:', error);
        if (!isOperationCancelled) showError('Error generating Excel file: ' + error.message);
    }
}

// =========================================================
// CHANGE 4: Deprecated apps download
// Stored globally so the deprecated download button can access them
// =========================================================
let _globalDeprecatedResults = [];
let _globalSelectedFields = [];
let _globalOriginalFioriIds = new Set();
let _globalReleaseId = '';

function generateDeprecatedExcel() {
    if (_globalDeprecatedResults.length === 0) {
        showError('No deprecated apps found to download');
        return;
    }
    try {
        const wb = buildExcelWorkbook(_globalDeprecatedResults, _globalSelectedFields, _globalOriginalFioriIds);
        if (wb) {
            XLSX.writeFile(wb, `Fiori_Apps_Deprecated_${_globalReleaseId}.xlsx`);
        } else {
            showError('No deprecated apps data available');
        }
    } catch (error) {
        console.error('Error generating deprecated Excel:', error);
        showError('Error generating deprecated Excel file: ' + error.message);
    }
}

// Setup deprecated download button
document.addEventListener('DOMContentLoaded', () => {
    const depBtn = document.getElementById('deprecatedDownloadButton');
    if (depBtn) {
        depBtn.addEventListener('click', generateDeprecatedExcel);
    }
});

// =========================================================
// Main Button Handler
// =========================================================
document.getElementById('customDownloadButton').addEventListener('click', async function () {
    isOperationCancelled = false;

    const releaseId = document.getElementById('releaseId').value.trim();
    const initialFioriIds = document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0);
    const selectedFields = getSelectedFields();

    if (initialFioriIds.length === 0 || !releaseId) {
        showError('Please enter both Fiori App IDs and Release ID');
        return;
    }
    if (selectedFields.length === 0) {
        showError('Please select at least one field to include in the Excel');
        return;
    }

    // CHANGE 2: Read the "remove no target mapping" checkbox
    const removeNoTargetMapping = document.getElementById('removeNoTargetMapping')?.checked || false;

    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');

    document.getElementById('loading').style.display = 'block';
    document.getElementById('results').style.display = 'none';
    document.getElementById('error').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';
    progressBar.style.width = '0%';
    progressText.textContent = 'Starting...';

    // Hide deprecated button while processing
    const depBtn = document.getElementById('deprecatedDownloadButton');
    if (depBtn) depBtn.style.display = 'none';

    try {
        const results = [];
        const processedApps = new Set();
        const queuedApps = [...initialFioriIds];
        const originalApps = new Set(initialFioriIds);
        let totalProcessed = 0;

        while (queuedApps.length > 0) {
            if (isOperationCancelled) break;

            const fioriId = queuedApps.shift();
            if (processedApps.has(fioriId)) continue;
            processedApps.add(fioriId);
            totalProcessed++;

            try {
                const totalExpectedApps = totalProcessed + queuedApps.length;
                const progressPercentage = Math.round((totalProcessed / Math.max(totalExpectedApps, 1)) * 100);
                progressBar.style.width = `${progressPercentage}%`;
                progressText.textContent = `Processing: ${totalProcessed} apps (${queuedApps.length} in queue)`;

                // CHANGE 1: Only fetch what's needed based on selected fields
                const {
                    appDetails, relatedApps, technicalNames, businessRoles,
                    bspNames, technicalCatalogs, spaces, pages, semanticObjects, semanticActions
                } = await fetchAppDataSelective(fioriId, releaseId, selectedFields);

                // CHANGE 2: For related apps, optionally filter out those with no target mappings
                const filteredRelatedApps = [];
                for (const relatedApp of relatedApps) {
                    const relatedFioriId = relatedApp.FioriId;
                    if (!processedApps.has(relatedFioriId) && !queuedApps.includes(relatedFioriId)) {
                        if (removeNoTargetMapping && !originalApps.has(relatedFioriId)) {
                            // Check if it has target mappings before queuing
                            const hasMappings = await hasTargetMappings(relatedFioriId, releaseId);
                            if (hasMappings) {
                                filteredRelatedApps.push(relatedApp);
                                queuedApps.push(relatedFioriId);
                            }
                            // else skip — no target mappings
                        } else {
                            filteredRelatedApps.push(relatedApp);
                            queuedApps.push(relatedFioriId);
                        }
                    } else {
                        filteredRelatedApps.push(relatedApp);
                    }
                }

                results.push({
                    fioriId,
                    status: 'Success',
                    isDeprecated: appDetails.isPublished === 'Deprecated',
                    isOriginalRequest: originalApps.has(fioriId),
                    appDetails,
                    technicalNames,
                    businessRoles,
                    bspNames,
                    technicalCatalogs,
                    spaces,
                    pages,
                    relatedApps: filteredRelatedApps,
                    semanticObjects,
                    semanticActions
                });
                updateProcessingResults(results);

            } catch (error) {
                results.push({
                    fioriId,
                    status: 'Error',
                    isOriginalRequest: originalApps.has(fioriId),
                    error: error.message
                });
                updateProcessingResults(results);
            }
        }

        if (!isOperationCancelled) {
            // CHANGE 3: Sort results by catalog name
            const sortedResults = sortResultsByCatalog(results, originalApps);

            // CHANGE 4: Separate deprecated original apps
            const deprecatedOriginalResults = results.filter(r =>
                r.status === 'Success' && r.isDeprecated && originalApps.has(r.fioriId)
            );

            // Store deprecated data globally for the download button
            _globalDeprecatedResults = deprecatedOriginalResults;
            _globalSelectedFields = [...selectedFields];
            _globalOriginalFioriIds = new Set(originalApps);
            _globalReleaseId = releaseId;

            // Generate main Excel (excludes deprecated)
            generateExcel(sortedResults, selectedFields, originalApps);

            // Show deprecated download button if there are deprecated apps
            if (deprecatedOriginalResults.length > 0 && depBtn) {
                depBtn.style.display = 'block';
                depBtn.textContent = `⬇️ Download Deprecated Apps (${deprecatedOriginalResults.length})`;
            }

            progressText.textContent = `Completed: Processed ${totalProcessed} apps`;
            progressBar.style.width = '100%';
            document.getElementById('loading').style.display = 'none';
            document.getElementById('results').style.display = 'block';
        }

    } catch (error) {
        showError('Error processing apps: ' + error.message);
        document.getElementById('loading').style.display = 'none';
    }
});
