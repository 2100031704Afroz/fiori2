document.getElementById('fetchButton').addEventListener('click', async function () {
    const fioriIds = document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0);
    const releaseId = document.getElementById('releaseId').value.trim();

    if (fioriIds.length === 0 || !releaseId) {
        showError('Please enter both Fiori App IDs and Release ID');
        return;
    }

    // Show loading, hide results and error
    document.getElementById('loading').style.display = 'block';
    document.getElementById('results').style.display = 'none';
    document.getElementById('error').style.display = 'none';
    document.getElementById('downloadButton').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';

    try {
        const results = [];
        const totalApps = fioriIds.length;
        let processedApps = 0;

        for (const fioriId of fioriIds) {
            try {
                // Update progress
                processedApps++;
                const progress = (processedApps / totalApps) * 100;
                document.getElementById('progressBar').style.width = `${progress}%`;
                document.getElementById('progressText').textContent = `Processing: ${Math.round(progress)}%`;

                // Fetch app details
                const appDetails = await fetchAppDetails(fioriId, releaseId);
                
                // Fetch all data for the app
                const [
                    technicalNames,
                    businessRoles,
                    bspNames,
                    technicalCatalogs,
                    spaces,
                    pages,
                    relatedApps,
                    semanticObjects
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

                // Fetch semantic actions
                const semanticActions = await retryFetch(() => fetchSemanticActions(fioriId, releaseId));

                // Add to results
                results.push({
                    fioriId,
                    status: 'Success',
                    isDeprecated: appDetails.isPublished === 'Deprecated',
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

                // Update processing results
                updateProcessingResults(results);
            } catch (error) {
                results.push({
                    fioriId,
                    status: 'Error',
                    error: error.message
                });
                updateProcessingResults(results);
            }
        }

        // Generate Excel
        generateExcel(results);

        // Hide loading, show results and download button
        document.getElementById('loading').style.display = 'none';
        document.getElementById('results').style.display = 'block';
        document.getElementById('downloadButton').style.display = 'block';
    } catch (error) {
        showError('Error processing apps: ' + error.message);
        document.getElementById('loading').style.display = 'none';
    }
});

function updateProcessingResults(results) {
    const container = document.getElementById('processingResults');
    container.innerHTML = '';

    results.forEach(result => {
        const div = document.createElement('div');
        div.className = 'data-item';
        
        let statusClass = result.status === 'Success' ? 'success' : 'error';
        let statusText = result.status === 'Success' ? 
            (result.isDeprecated ? 'Success (Deprecated)' : 'Success') : 
            'Error';

        let html = `
            <p><strong>${result.fioriId}</strong> - <span class="${statusClass}">${statusText}</span></p>
        `;

        if (result.status === 'Error') {
            html += `<p class="data-item-details">Error: ${result.error}</p>`;
        } else if (result.isDeprecated) {
            html += `<p class="data-item-details">⚠️ This application is deprecated in the selected release</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

function generateExcel(results) {
    // Create workbook
    const wb = XLSX.utils.book_new();

    // Filter out deprecated apps
    const validResults = results.filter(result => result.status === 'Success' && !result.isDeprecated);

    if (validResults.length === 0) {
        showError('No valid apps found to generate Excel');
        return;
    }

    // Create sets to store unique values
    const uniqueBusinessRoles = new Set();
    const uniqueODataServices = new Set();
    const uniqueTechnicalCatalogs = new Set();
    const uniqueBSPNames = new Set();
    const uniqueSpaces = new Set();
    const uniquePages = new Set();
    const uniqueRelatedApps = new Set();
    const uniqueSemanticObjects = new Set();
    const uniqueSemanticActions = new Set();

    // Collect all unique values
    validResults.forEach(result => {
        result.businessRoles.forEach(role => uniqueBusinessRoles.add(role.BusinessRoleName));
        result.technicalNames.forEach(service => uniqueODataServices.add(service.TechnicalName));
        result.technicalCatalogs.forEach(catalog => uniqueTechnicalCatalogs.add(catalog.TechincalCatalog));
        // Add BSP name from app details if it exists
        if (result.appDetails.BSPName) {
            uniqueBSPNames.add(result.appDetails.BSPName);
        }
        // Add BSP names from BSP names data
        result.bspNames.forEach(bsp => {
            if (bsp.BSPName) {
                uniqueBSPNames.add(bsp.BSPName);
            }
        });
        result.spaces.forEach(space => uniqueSpaces.add(space.SpaceName));
        result.pages.forEach(page => uniquePages.add(page.PageName));
        result.relatedApps.forEach(app => uniqueRelatedApps.add(app.FioriId));
        result.semanticObjects.forEach(obj => {
            uniqueSemanticObjects.add(obj.SemanticObject);
            if (result.semanticActions) {
                result.semanticActions.forEach(sa => {
                    if (sa.SemanticObject && sa.SemanticAction) {
                        uniqueSemanticActions.add(`${sa.SemanticObject}:${sa.SemanticAction}`);
                    }
                });
            }
        });
    });

    // Create data for the single sheet
    const excelData = validResults.map(result => {
        const row = {
            'Fiori ID': result.fioriId,
            'App Title': result.appDetails.Title || result.appDetails.AppName,
            'Application Type': result.appDetails.ApplicationType,
            'UI Technology': result.appDetails.UITechnology,
            'Application Component': result.appDetails.ApplicationComponent,
            'BSP Name': result.appDetails.BSPName || '',
            'UI5 Component ID': result.appDetails.SAPUI5ComponentId,
            'Business Roles': result.businessRoles.map(role => role.BusinessRoleName).join('\n'),
            'OData Services': result.technicalNames.map(service => service.TechnicalName).join('\n'),
            'Technical Catalogs': result.technicalCatalogs.map(catalog => catalog.TechincalCatalog).join('\n'),
            'Spaces': result.spaces.map(space => space.SpaceName).join('\n'),
            'Pages': result.pages.map(page => page.PageName).join('\n'),
            'Related Apps': result.relatedApps.map(app => app.FioriId).join('\n'),
            'Semantic Objects': result.semanticObjects.map(obj => obj.SemanticObject).join('\n'),
            'Semantic Actions': result.semanticActions ? result.semanticActions.map(sa => `${sa.SemanticObject}:${sa.SemanticAction}`).join('\n') : ''
        };

        return row;
    });

    // Add consolidated columns
    excelData.push({
        'Fiori ID': 'CONSOLIDATED',
        'App Title': '',
        'Application Type': '',
        'UI Technology': '',
        'Application Component': '',
        'BSP Name': Array.from(uniqueBSPNames).sort().join('\n'),
        'UI5 Component ID': '',
        'Business Roles': Array.from(uniqueBusinessRoles).sort().join('\n'),
        'OData Services': Array.from(uniqueODataServices).sort().join('\n'),
        'Technical Catalogs': Array.from(uniqueTechnicalCatalogs).sort().join('\n'),
        'Spaces': Array.from(uniqueSpaces).sort().join('\n'),
        'Pages': Array.from(uniquePages).sort().join('\n'),
        'Related Apps': Array.from(uniqueRelatedApps).sort().join('\n'),
        'Semantic Objects': Array.from(uniqueSemanticObjects).sort().join('\n'),
        'Semantic Actions': Array.from(uniqueSemanticActions).sort().join('\n')
    });

    // Create the worksheet
    const ws = XLSX.utils.json_to_sheet(excelData);

    // Set column widths
    const colWidths = {
        'A': 15, // Fiori ID
        'B': 40, // App Title
        'C': 20, // Application Type
        'D': 20, // UI Technology
        'E': 30, // Application Component
        'F': 20, // BSP Name
        'G': 30, // UI5 Component ID
        'H': 40, // Business Roles
        'I': 40, // OData Services
        'J': 40, // Technical Catalogs
        'K': 40, // Spaces
        'L': 40, // Pages
        'M': 40, // Related Apps
        'N': 40,  // Semantic Objects
        'O': 40   // Semantic Actions
    };

    ws['!cols'] = Object.values(colWidths).map(width => ({ wch: width }));

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Fiori Apps Data');

    // Save the workbook
    const releaseId = document.getElementById('releaseId').value.trim();
    XLSX.writeFile(wb, `Fiori_Apps_Data_${releaseId}.xlsx`);
}

// Add event listener for download button
document.getElementById('downloadButton').addEventListener('click', function() {
    const releaseId = document.getElementById('releaseId').value.trim();
    const fioriIds = document.getElementById('fioriIds').value.trim().split(' ').filter(id => id.length > 0);
    
    // Trigger the fetch process again to generate fresh Excel
    document.getElementById('fetchButton').click();
});

async function retryFetch(fetchFn, maxRetries = 3) {
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await fetchFn();
        } catch (error) {
            if (i === maxRetries - 1) throw error;
            // Wait before retrying (exponential backoff)
            await new Promise(resolve => setTimeout(resolve, Math.pow(2, i) * 1000));
        }
    }
}

// Enhance error display
function showError(message) {
    const errorElement = document.getElementById('error');
    errorElement.textContent = message;
    errorElement.style.display = 'block';
    // Auto-hide error after 5 seconds
    setTimeout(() => {
        errorElement.style.display = 'none';
    }, 5000);
}

async function fetchTechnicalCatalogs(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/SplitTechnicalCatalogs?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

function displayTechnicalCatalogs(technicalCatalogs) {
    const container = document.getElementById('technicalCatalogs');
    const textContainer = document.getElementById('technicalCatalogsText');
    container.innerHTML = '';

    if (technicalCatalogs.length === 0) {
        container.innerHTML = '<p>No Technical Catalogs found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the catalog names for easy copying
    const namesOnly = technicalCatalogs.map(item => item.TechincalCatalog);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyTechnicalCatalogs').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyTechnicalCatalogsSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    technicalCatalogs.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.TechincalCatalog}</strong></p>`;

        if (item.TechincalCatalogDescription) {
            html += `<p class="data-item-details">Description: ${item.TechincalCatalogDescription}</p>`;
        }

        if (item.SystemAlias) {
            html += `<p class="data-item-details">System Alias: ${item.SystemAlias}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

// Add this function to fetch semantic objects
async function fetchSemanticObjects(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/SplitAdditionalIntents?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

// Add this function to display semantic objects
function displaySemanticObjects(semanticObjects) {
    const container = document.getElementById('semanticObjects');
    const textContainer = document.getElementById('semanticObjectsText');
    container.innerHTML = '';

    if (semanticObjects.length === 0) {
        container.innerHTML = '<p>No Semantic Objects found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the semantic object names for easy copying
    const objectsOnly = semanticObjects.map(item => item.SemanticObject);
    // Remove duplicates
    const uniqueObjects = [...new Set(objectsOnly)];
    textContainer.textContent = uniqueObjects.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copySemanticObjects').addEventListener('click', async () => {
        const success = await copyToClipboard(uniqueObjects.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copySemanticObjectsSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    semanticObjects.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.SemanticObject}</strong></p>`;

        if (item.SemanticAction) {
            html += `<p class="data-item-details">Action: ${item.SemanticAction}</p>`;
        }

        if (item.MappingSignatureKeyVal) {
            html += `<p class="data-item-details">Parameters: ${item.MappingSignatureKeyVal}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

// Add this function to fetch related apps
async function fetchRelatedApps(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/Related_Apps?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

// Add this function to display related apps
function displayRelatedApps(relatedApps) {
    const container = document.getElementById('relatedApps');
    const textContainer = document.getElementById('relatedAppsText');
    container.innerHTML = '';

    if (relatedApps.length === 0) {
        container.innerHTML = '<p>No Related Apps found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the Fiori IDs for easy copying (excluding App Names as requested)
    const idsOnly = relatedApps.map(item => item.FioriId);
    textContainer.textContent = idsOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyRelatedApps').addEventListener('click', async () => {
        const success = await copyToClipboard(idsOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyRelatedAppsSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items - showing both App Name and Fiori ID
    relatedApps.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.FioriId}</strong></p>`;

        if (item.AppName) {
            html += `<p class="data-item-details">App Name: ${item.AppName}</p>`;
        }

        if (item.relationType) {
            html += `<p class="data-item-details">Relation Type: ${item.relationType}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

async function fetchSpaces(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/SplitSpace?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

function displaySpaces(spaces) {
    const container = document.getElementById('spaces');
    const textContainer = document.getElementById('spacesText');
    container.innerHTML = '';

    if (spaces.length === 0) {
        container.innerHTML = '<p>No Spaces found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the space names for easy copying
    const namesOnly = spaces.map(item => item.SpaceName);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copySpaces').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copySpacesSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    spaces.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.SpaceName}</strong></p>`;

        if (item.SpaceTitle) {
            html += `<p class="data-item-details">Title: ${item.SpaceTitle}</p>`;
        }

        if (item.SpaceDescription) {
            html += `<p class="data-item-details">Description: ${item.SpaceDescription}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

async function fetchPages(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/SplitPage?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

function displayPages(pages) {
    const container = document.getElementById('pages');
    const textContainer = document.getElementById('pagesText');
    container.innerHTML = '';

    if (pages.length === 0) {
        container.innerHTML = '<p>No Pages found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the page names for easy copying
    const namesOnly = pages.map(item => item.PageName);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyPages').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyPagesSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    pages.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.PageName}</strong></p>`;

        if (item.PageTitle) {
            html += `<p class="data-item-details">Title: ${item.PageTitle}</p>`;
        }

        if (item.PageDescription) {
            html += `<p class="data-item-details">Description: ${item.PageDescription}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

async function fetchAppDetails(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(fioriId='${fioriId}',releaseId='${releaseId}',inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN')?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d;
}

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
            <div class="app-meta-item">
                <span class="app-meta-label">Fiori ID:</span>
                <span>${appData.fioriId || 'N/A'}</span>
            </div>
            <div class="app-meta-item">
                <span class="app-meta-label">Release:</span>
                <span>${appData.ReleaseName || appData.releaseId || 'N/A'}</span>
            </div>
            <div class="app-meta-item">
                <span class="app-meta-label">Application Type:</span>
                <span>${appData.ApplicationType || 'N/A'}</span>
            </div>
            <div class="app-meta-item">
                <span class="app-meta-label">UI Technology:</span>
                <span>${appData.UITechnology || 'N/A'}</span>
            </div>
        </div>
    `;

    // Check if app is deprecated
    if (appData.isPublished === 'Deprecated') {
        html += `
            <div class="deprecated-warning">
                <span class="deprecated-icon">⚠️</span>
                This application is marked as DEPRECATED in this release
            </div>
        `;
    }

    // Add Application Component if available
    if (appData.ApplicationComponent) {
        html += `
            <div class="app-meta-item">
                <span class="app-meta-label">Application Component:</span>
                <span>${appData.ApplicationComponent} (${appData.ApplicationComponentText || ''})</span>
            </div>
        `;
    }

    // Add BSP Name if available
    if (appData.BSPName) {
        html += `
            <div class="app-meta-item">
                <span class="app-meta-label">BSP Application:</span>
                <span>${appData.BSPName}</span>
            </div>
        `;
    }

    // Add UI5 Component ID if available
    if (appData.SAPUI5ComponentId) {
        html += `
            <div class="app-meta-item">
                <span class="app-meta-label">SAPUI5 Component ID:</span>
                <span>${appData.SAPUI5ComponentId}</span>
            </div>
        `;
    }

    container.innerHTML = html;
    appInfoSection.style.display = 'block';
}

async function fetchTechnicalNames(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/ODataServices?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

async function fetchBusinessRoleNames(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(fioriId='${fioriId}',releaseId='${releaseId}',inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN')/SplitBusinessRole?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

async function fetchBSPNames(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/ICFNodes?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.d && data.d.results ? data.d.results : [];
}

// Helper function to copy text to clipboard
async function copyToClipboard(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch (err) {
        console.error('Failed to copy: ', err);
        return false;
    }
}

function displayTechnicalNames(technicalNames) {
    const container = document.getElementById('technicalNames');
    const textContainer = document.getElementById('technicalNamesText');
    container.innerHTML = '';

    if (technicalNames.length === 0) {
        container.innerHTML = '<p>No Technical Names found</p>';
        textContainer.textContent = '';
        return;
    }

    // Create a list of just the names for easy copying
    const namesOnly = technicalNames.map(item => item.TechnicalName);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyTechnicalNames').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyTechnicalSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    technicalNames.forEach(item => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${item.TechnicalName}</strong></p>`;

        if (item.NameSpace) {
            html += `<p class="data-item-details">Namespace: ${item.NameSpace}</p>`;
        }

        if (item.SoftwareComponentName) {
            html += `<p class="data-item-details">Software Component: ${item.SoftwareComponentName}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

function displayBusinessRoles(businessRoles) {
    const container = document.getElementById('businessRoles');
    const textContainer = document.getElementById('businessRolesText');
    container.innerHTML = '';

    if (businessRoles.length === 0) {
        container.innerHTML = '<p>No Business Roles found</p>';
        textContainer.textContent = '';
        return;
    }

    // Sort to show leading roles first
    businessRoles.sort((a, b) => {
        if (a.isLeading === "X" && b.isLeading !== "X") return -1;
        if (a.isLeading !== "X" && b.isLeading === "X") return 1;
        return 0;
    });

    // Create a list of just the names for easy copying
    const namesOnly = businessRoles.map(role => role.BusinessRoleName);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyBusinessRoles').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyBusinessSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    businessRoles.forEach(role => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${role.BusinessRoleName}</strong>${role.isLeading === "X" ? ' (Leading Role)' : ''}</p>`;

        if (role.BusinessRoleDescription) {
            html += `<p class="data-item-details">Description: ${role.BusinessRoleDescription}</p>`;
        }

        if (role.RoleID) {
            html += `<p class="data-item-details">Role ID: ${role.RoleID}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

function displayBSPNames(bspNames) {
    const container = document.getElementById('bspNames');
    const textContainer = document.getElementById('bspNamesText');
    container.innerHTML = '';

    if (bspNames.length === 0) {
        container.innerHTML = '<p>No BSP Names found</p>';
        textContainer.textContent = '';
        return;
    }

    // Sort to show main (non-additional) BSPs first
    bspNames.sort((a, b) => a.isAdditional - b.isAdditional);

    // Create a list of just the names for easy copying
    const namesOnly = bspNames.map(bsp => bsp.BSPName);
    textContainer.textContent = namesOnly.join('\n');
    textContainer.style.display = 'block';

    // Add copy button functionality
    document.getElementById('copyBSPNames').addEventListener('click', async () => {
        const success = await copyToClipboard(namesOnly.join('\n'));
        if (success) {
            const successIndicator = document.getElementById('copyBSPSuccess');
            successIndicator.style.display = 'inline';
            setTimeout(() => { successIndicator.style.display = 'none'; }, 2000);
        }
    });

    // Display detailed items
    bspNames.forEach(bsp => {
        const div = document.createElement('div');
        div.className = 'data-item';

        let html = `<p><strong>${bsp.BSPName}</strong>${bsp.isAdditional === 0 ? ' (Main BSP)' : ''}</p>`;

        if (bsp.AppName) {
            html += `<p class="data-item-details">App Name: ${bsp.AppName}</p>`;
        }

        if (bsp.BSPApplicationURL) {
            html += `<p class="data-item-details">URL: ${bsp.BSPApplicationURL}</p>`;
        }

        if (bsp.SAPUI5ComponentId) {
            html += `<p class="data-item-details">UI5 Component ID: ${bsp.SAPUI5ComponentId}</p>`;
        }

        div.innerHTML = html;
        container.appendChild(div);
    });
}

document.getElementById('fioriIds').addEventListener('keypress', handleEnterKey);
document.getElementById('releaseId').addEventListener('keypress', handleEnterKey);

// Add event listeners to convert input to uppercase
document.getElementById('fioriIds').addEventListener('input', function (e) {
    e.target.value = e.target.value.toUpperCase();
});
document.getElementById('releaseId').addEventListener('input', function (e) {
    e.target.value = e.target.value.toUpperCase();
});

function handleEnterKey(event) {
    if (event.key === 'Enter') {
        event.preventDefault(); // Prevent default form submission
        if (document.getElementById('fioriIds').value === '' ||
            document.getElementById('releaseId').value === '') {
            showError('Please enter both Fiori App IDs and Release ID');
            return;
        }

        const fioriIds = document.getElementById('fioriIds').value.trim().split(' ');
        const releaseId = document.getElementById('releaseId').value.trim();

        // Only trigger fetch if both fields have values
        if (fioriIds.length > 0 && releaseId) {
            document.getElementById('fetchButton').click();
        }
    }
}

// Add this function to fetch semantic actions
async function fetchSemanticActions(fioriId, releaseId) {
    const url = `https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Details(inpfioriId='${fioriId}',inpreleaseId='${releaseId}',inpLanguage='EN',fioriId='${fioriId}',releaseId='${releaseId}')/SplitAdditionalIntents?$format=json`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    if (data.d && data.d.results) {
        // Map to array of { SemanticObject, SemanticAction }
        return data.d.results.map(item => ({
            SemanticObject: item.SemanticObject,
            SemanticAction: item.SemanticAction
        }));
    }
    return [];
}
