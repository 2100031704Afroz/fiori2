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
    container.innerHTML = '';

    if (!data?.length) {
        container.innerHTML = `<p>No ${containerId} found</p>`;
        textEl.textContent = '';
        return;
    }

    const names = [...new Set(data.map(nameFn))];
    textEl.textContent = names.join('\n');
    textEl.style.display = 'block';

    document.getElementById(copyButtonId).onclick = async () => {
        if (await copyToClipboard(names.join('\n'))) {
            const success = document.getElementById(successId);
            success.style.display = 'inline';
            setTimeout(() => (success.style.display = 'none'), 2000);
        }
    };

    data.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';
        div.innerHTML = detailsFn(item);
        container.appendChild(div);
    });
}

// --- Display Functions ---
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
    if (!appData) {
        container.innerHTML = '<p>No app information available</p>';
        return;
    }
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
    appInfoSection.style.display = 'block';
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

// --- Excel Generation ---
function generateExcel(results, selectedFields) {
    const wb = XLSX.utils.book_new();
    const validResults = results.filter(result => result.status === 'Success' && !result.isDeprecated);
    if (validResults.length === 0) {
        showError('No valid apps found to generate Excel');
        return;
    }
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
            if (result.semanticActions) {
                result.semanticActions.forEach(sa => {
                    if (sa.SemanticObject && sa.SemanticAction) 
                        uniqueSets.semanticActions.add(`${sa.SemanticObject}:${sa.SemanticAction}`);
                });
            }
        });
    });
    // Field mapping for Excel
    const fieldMapping = {
        'Fiori ID':           (result) => result.fioriId,
        'App Title':          (result) => result.appDetails.Title || result.appDetails.AppName,
        'Application Type':   (result) => result.appDetails.ApplicationType,
        'UI Technology':      (result) => result.appDetails.UITechnology,
        'Application Component': (result) => result.appDetails.ApplicationComponent,
        'BSP Name':           (result) => {
            let bspNamesList = [];
            if (result.appDetails.BSPName) bspNamesList.push(result.appDetails.BSPName);
            if (result.bspNames) result.bspNames.forEach(bsp => {
                if (bsp.BSPName && !bspNamesList.includes(bsp.BSPName)) bspNamesList.push(bsp.BSPName);
            });
            return bspNamesList.join('\n');
        },
        'UI5 Component ID':   (result) => result.appDetails.SAPUI5ComponentId,
        'Business Roles':     (result) => result.businessRoles.map(role => role.BusinessRoleName).join('\n'),
        'OData Services':     (result) => result.technicalNames.map(service => service.TechnicalName).join('\n'),
        'Technical Catalogs': (result) => result.technicalCatalogs.map(catalog => catalog.TechincalCatalog).join('\n'),
        'Spaces':             (result) => result.spaces.map(space => space.SpaceName).join('\n'),
        'Pages':              (result) => result.pages.map(page => page.PageName).join('\n'),
        'Related Apps':       (result) => result.relatedApps.map(app => app.FioriId).join('\n'),
        'Semantic Objects':   (result) => result.semanticObjects.map(obj => obj.SemanticObject).join('\n'),
        'Semantic Actions':   (result) => result.semanticActions ? result.semanticActions.map(sa => `${sa.SemanticObject}:${sa.SemanticAction}`).join('\n') : ''
    };
    // Consolidated mapping
    const consolidatedMapping = {
        'Fiori ID': 'CONSOLIDATED',
        'App Type': '-',
        'BSP Name': Array.from(uniqueSets.bspNames).sort().join('\n'),
        'Business Roles': Array.from(uniqueSets.businessRoles).sort().join('\n'),
        'OData Services': Array.from(uniqueSets.odataServices).sort().join('\n'),
        'Technical Catalogs': Array.from(uniqueSets.technicalCatalogs).sort().join('\n'),
        'Spaces': Array.from(uniqueSets.spaces).sort().join('\n'),
        'Pages': Array.from(uniqueSets.pages).sort().join('\n'),
        'Related Apps': Array.from(uniqueSets.relatedApps).sort().join('\n'),
        'Semantic Objects': Array.from(uniqueSets.semanticObjects).sort().join('\n'),
        'Semantic Actions': Array.from(uniqueSets.semanticActions).sort().join('\n')
    };
    // Add App Type field if missing
    if (!selectedFields.includes('App Type')) selectedFields.push('App Type');
    // Excel data
    const originalFioriIds = new Set(document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0));
    const excelData = validResults.map(result => {
        const row = {};
        selectedFields.forEach(field => {
            if (field === 'App Type') {
                row[field] = originalFioriIds.has(result.fioriId) ? '✓' : '✗';
            } else {
                row[field] = fieldMapping[field] ? fieldMapping[field](result) : '';
            }
        });
        return row;
    });
    // Consolidated row
    const consolidatedRow = {};
    selectedFields.forEach(field => { consolidatedRow[field] = consolidatedMapping[field] || ''; });
    excelData.push(consolidatedRow);
    // Create worksheet/Excel
    const ws = XLSX.utils.json_to_sheet(excelData);
    const colWidthsMap = {
        'Fiori ID': 15, 'App Title': 40, 'Application Type': 20, 'UI Technology': 20,
        'Application Component': 30, 'BSP Name': 20, 'UI5 Component ID': 30,
        'Business Roles': 40, 'OData Services': 40, 'Technical Catalogs': 40,
        'App Type': 10, 'Spaces': 40, 'Pages': 40, 'Related Apps': 40,
        'Semantic Objects': 40, 'Semantic Actions': 40
    };
    ws['!cols'] = selectedFields.map(field => ({ wch: colWidthsMap[field] || 20 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Fiori Apps Data');
    const releaseId = document.getElementById('releaseId').value.trim();
    XLSX.writeFile(wb, `Fiori_Apps_Data_${releaseId}.xlsx`);
}

// --- Button Handler and App Processing ---
function generateExcel(results, selectedFields) {
    // Check if operation was cancelled before generating Excel
    if (isOperationCancelled) {
        console.log('Excel generation cancelled');
        return;
    }
    
    const wb = XLSX.utils.book_new();
    const validResults = results.filter(result => result.status === 'Success' && !result.isDeprecated);

    if (validResults.length === 0) {
        showError('No valid apps found to generate Excel');
        return;
    }

    try {
        // Create sets to store unique values
        const uniqueSets = {
            businessRoles: new Set(),
            odataServices: new Set(),
            technicalCatalogs: new Set(),
            bspNames: new Set(),
            spaces: new Set(),
            pages: new Set(),
            relatedApps: new Set(),
            semanticObjects: new Set(),
            semanticActions: new Set()
        };

        // Collect all unique values
        validResults.forEach(result => {
            result.businessRoles.forEach(role => uniqueSets.businessRoles.add(role.BusinessRoleName));
            result.technicalNames.forEach(service => uniqueSets.odataServices.add(service.TechnicalName));
            result.technicalCatalogs.forEach(catalog => uniqueSets.technicalCatalogs.add(catalog.TechincalCatalog));
            if (result.appDetails.BSPName) uniqueSets.bspNames.add(result.appDetails.BSPName);
            result.bspNames.forEach(bsp => {
                if (bsp.BSPName) uniqueSets.bspNames.add(bsp.BSPName);
            });
            result.spaces.forEach(space => uniqueSets.spaces.add(space.SpaceName));
            result.pages.forEach(page => uniqueSets.pages.add(page.PageName));
            result.relatedApps.forEach(app => uniqueSets.relatedApps.add(app.FioriId));
            result.semanticObjects.forEach(obj => {
                uniqueSets.semanticObjects.add(obj.SemanticObject);
                if (result.semanticActions) {
                    result.semanticActions.forEach(sa => {
                        if (sa.SemanticObject && sa.SemanticAction) {
                            uniqueSets.semanticActions.add(`${sa.SemanticObject}:${sa.SemanticAction}`);
                        }
                    });
                }
            });
        });

        // Add App Type field to selectedFields if not present
        if (!selectedFields.includes('App Type')) {
            selectedFields.push('App Type');
        }

        // Create data for the single sheet
        const originalFioriIds = new Set(document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0));
        const excelData = validResults.map(result => {
            const row = {};
            selectedFields.forEach(field => {
                if (field === 'App Type') {
                    row[field] = originalFioriIds.has(result.fioriId) ? '✓' : '✗';
                } else {
                    switch(field) {
                        case 'Fiori ID': row[field] = result.fioriId; break;
                        case 'App Title': row[field] = result.appDetails.Title || result.appDetails.AppName; break;
                        case 'Application Type': row[field] = result.appDetails.ApplicationType; break;
                        case 'UI Technology': row[field] = result.appDetails.UITechnology; break;
                        case 'Application Component': row[field] = result.appDetails.ApplicationComponent; break;
                        case 'BSP Name': 
                            let bspNames = [];
                            if (result.appDetails.BSPName) bspNames.push(result.appDetails.BSPName);
                            result.bspNames.forEach(bsp => {
                                if (bsp.BSPName && !bspNames.includes(bsp.BSPName)) {
                                    bspNames.push(bsp.BSPName);
                                }
                            });
                            row[field] = bspNames.join('\n');
                            break;
                        case 'UI5 Component ID': row[field] = result.appDetails.SAPUI5ComponentId; break;
                        case 'Business Roles': row[field] = result.businessRoles.map(role => role.BusinessRoleName).join('\n'); break;
                        case 'OData Services': row[field] = result.technicalNames.map(service => service.TechnicalName).join('\n'); break;
                        case 'Technical Catalogs': row[field] = result.technicalCatalogs.map(catalog => catalog.TechincalCatalog).join('\n'); break;
                        case 'Spaces': row[field] = result.spaces.map(space => space.SpaceName).join('\n'); break;
                        case 'Pages': row[field] = result.pages.map(page => page.PageName).join('\n'); break;
                        case 'Related Apps': row[field] = result.relatedApps.map(app => app.FioriId).join('\n'); break;
                        case 'Semantic Objects': row[field] = result.semanticObjects.map(obj => obj.SemanticObject).join('\n'); break;
                        case 'Semantic Actions': row[field] = result.semanticActions ? result.semanticActions.map(sa => `${sa.SemanticObject}:${sa.SemanticAction}`).join('\n') : ''; break;
                    }
                }
            });
            return row;
        });

        // Add consolidated row
        const consolidatedRow = {};
        selectedFields.forEach(field => {
            switch(field) {
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

        // Create worksheet
        const ws = XLSX.utils.json_to_sheet(excelData);

        // Set column widths
        const colWidthsMap = {
            'Fiori ID': 15, 'App Title': 40, 'Application Type': 20, 'UI Technology': 20,
            'Application Component': 30, 'BSP Name': 20, 'UI5 Component ID': 30,
            'Business Roles': 40, 'OData Services': 40, 'Technical Catalogs': 40,
            'App Type': 10, 'Spaces': 40, 'Pages': 40, 'Related Apps': 40,
            'Semantic Objects': 40, 'Semantic Actions': 40
        };
        ws['!cols'] = selectedFields.map(field => ({ wch: colWidthsMap[field] || 20 }));

        // Add the worksheet to workbook and save
        if (!isOperationCancelled) {
            XLSX.utils.book_append_sheet(wb, ws, 'Fiori Apps Data');
            const releaseId = document.getElementById('releaseId').value.trim();
            XLSX.writeFile(wb, `Fiori_Apps_Data_${releaseId}.xlsx`);
        }
    } catch (error) {
        console.error('Error generating Excel:', error);
        if (!isOperationCancelled) {
            showError('Error generating Excel file: ' + error.message);
        }
    }
}

document.getElementById('customDownloadButton').addEventListener('click', async function() {
    // Reset the cancelled flag and setup cancel button when starting a new operation
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

    // Reset UI elements
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    
    document.getElementById('loading').style.display = 'block';
    document.getElementById('results').style.display = 'none';
    document.getElementById('error').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';
    
    // Reset progress bar and text
    progressBar.style.width = '0%';
    progressText.textContent = 'Starting...';

    try {
        const results = [];
        const processedApps = new Set();
        const queuedApps = [...initialFioriIds];
        const originalApps = new Set(initialFioriIds);
        let totalProcessed = 0;

        while (queuedApps.length > 0) {
            const fioriId = queuedApps.shift();
            if (processedApps.has(fioriId)) continue;
            processedApps.add(fioriId);
            totalProcessed++;
            try {
                // Update progress
                const totalExpectedApps = totalProcessed + queuedApps.length;
                const progressPercentage = Math.round((totalProcessed / Math.max(totalExpectedApps, 1)) * 100);
                document.getElementById('progressBar').style.width = `${progressPercentage}%`;
                document.getElementById('progressText').textContent = `Processing: ${totalProcessed} apps (${queuedApps.length} in queue)`;

                // Fetch app details
                const appDetails = await fetchAppDetails(fioriId, releaseId);
                const [
                    technicalNames, businessRoles, bspNames,
                    technicalCatalogs, spaces, pages, relatedApps, semanticObjects
                ] = await Promise.all([
                    retryFetch(() => fetchTechnicalNames(fioriId, releaseId)),
                    retryFetch(() => fetchBusinessRoleNames(fioriId, releaseId)),
                    retryFetch(() => fetchBSPNames(fioriId, releaseId)),
                    retryFetch(() => fetchTechnicalCatalogs(fioriId, releaseId)),
                    retryFetch(() => fetchSpaces(fioriId, releaseId)),
                    retryFetch(() => fetchPages(fioriId, releaseId)),
                    retryFetch(() => fetchRelatedApps(fioriId, releaseId)),
                    retryFetch(() => fetchSemanticObjects(fioriId, releaseId))
                ]);
                const semanticActions = await retryFetch(() => fetchSemanticActions(fioriId, releaseId));
                relatedApps.forEach(relatedApp => {
                    const relatedFioriId = relatedApp.FioriId;
                    if (!processedApps.has(relatedFioriId) && !queuedApps.includes(relatedFioriId)) {
                        queuedApps.push(relatedFioriId);
                    }
                });
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
                    relatedApps,
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
            generateExcel(results, selectedFields);
            document.getElementById('progressText').textContent = `Completed: Processed ${totalProcessed} apps`;
            document.getElementById('progressBar').style.width = '100%';
            document.getElementById('loading').style.display = 'none';
            document.getElementById('results').style.display = 'block';
        }

    } catch (error) {
        showError('Error processing apps: ' + error.message);
        document.getElementById('loading').style.display = 'none';
    }
});
