const GET_PROJECTS_URL = "https://default062a8e8e449048f39ee3b309e2cfa4.ad.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/405ddbd55c224c9ebe1d2bc5b85a6597/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=P9Tu-84M5_ZRI2lryh6GQTPq9erJ9yTd9JNk0CVZli4";
const EXECUTE_ARCHIVE_URL = "https://emailpdfbackend.vercel.app/api/archive";
const STORAGE_KEY = "recentProjectsList";

let currentDirectoryPath = "";
let selectedTargetPath = "";

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("System Ready: Initializing UI...");
        renderQuickArchiveUI(); // Render on startup
        fetchDirectories("", 2);
        document.getElementById("executeBtn").onclick = () => triggerArchivePipeline(selectedTargetPath);
    }
});

const isProjectFolder = (folderName) => /^\d{5}/.test(folderName);

// 1. Unified Render Function
function renderQuickArchiveUI() {
    const quickSection = document.getElementById("quickSection");
    const quickList = document.getElementById("quickList");
    
    let recents = [];
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        recents = stored ? JSON.parse(stored) : [];
    } catch(e) {
        console.error("UI Render Error:", e);
        recents = [];
    }

    if (recents && recents.length > 0) {
        quickList.innerHTML = "";
        recents.forEach(pathData => {
            const btn = document.createElement("button");
            btn.className = "quick-btn";
            const displayName = pathData.split('/').pop();
            btn.innerHTML = `<span>${displayName}</span><span>&rarr;</span>`;
            btn.onclick = () => triggerArchivePipeline(pathData);
            quickList.appendChild(btn);
        });
        quickSection.style.display = "block";
    } else {
        quickSection.style.display = "none";
    }
}

// 2. Unified Storage Logic
function pushToRecentsStack(fullPath) {
    if (!fullPath) return;

    let recents = [];
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        recents = stored ? JSON.parse(stored) : [];
    } catch(e) {
        recents = [];
    }

    // Filter out existing to avoid duplicates, add to front
    recents = recents.filter(item => item !== fullPath);
    recents.unshift(fullPath);
    
    // Keep only last 10
    if (recents.length > 10) recents = recents.slice(0, 10);
    
    localStorage.setItem(STORAGE_KEY, JSON.stringify(recents));
    
    // Trigger immediate UI refresh
    renderQuickArchiveUI();
}

async function fetchDirectories(path, expectedSegments) {
    const container = document.getElementById("directoryBrowser");
    const pathDisplay = document.getElementById("currentPathDisplay");
    
    container.innerHTML = '<div style="padding:12px;color:#666;">Loading...</div>';
    document.getElementById("executeBtn").disabled = true;
    selectedTargetPath = "";
    
    pathDisplay.innerText = path === "" ? "/Project_Folder/" : `/Project_Folder/${path}/`;

    try {
        const response = await fetch(GET_PROJECTS_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path, expectedSegments })
        });
        
        if (!response.ok) throw new Error("API Connection Failed");
        const folders = await response.json();
        
        container.innerHTML = "";
        
        if (path !== "") {
            const backDiv = document.createElement('div');
            backDiv.className = "folder-item back-btn";
            backDiv.innerHTML = "<span class='icon'>&#8629;</span> Go Back";
            backDiv.onclick = () => {
                let segments = path.split('/');
                segments.pop();
                currentDirectoryPath = segments.join('/');
                fetchDirectories(currentDirectoryPath, expectedSegments - 1);
            };
            container.appendChild(backDiv);
        }

        folders.forEach(name => {
            const div = document.createElement('div');
            div.className = "folder-item";
            if (isProjectFolder(name)) {
                div.innerHTML = `<span class='icon'>&#128194;</span> ${name}`;
                div.onclick = () => {
                    document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('selected'));
                    div.classList.add('selected');
                    selectedTargetPath = currentDirectoryPath === "" ? name : `${currentDirectoryPath}/${name}`;
                    document.getElementById("executeBtn").disabled = false;
                };
            } else {
                div.innerHTML = `<span class='icon'>&#128193;</span> ${name}`;
                div.onclick = () => {
                    currentDirectoryPath = currentDirectoryPath === "" ? name : `${currentDirectoryPath}/${name}`;
                    fetchDirectories(currentDirectoryPath, expectedSegments + 1);
                };
            }
            container.appendChild(div);
        });
    } catch (error) {
        container.innerHTML = '<div style="padding:12px;color:red;">Error loading directory.</div>';
        console.error(error);
    }
}

// 3. Updated Pipeline with UI callback
async function triggerArchivePipeline(targetFullPath) {
    const status = document.getElementById("status");
    const item = Office.context.mailbox.item;

    if (!item || !item.itemId) return;

    status.innerText = "Archiving...";
    
    try {
        const convertedRestId = Office.context.mailbox.convertToRestId(
            item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );

        const response = await fetch(EXECUTE_ARCHIVE_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ itemId: convertedRestId, targetProject: targetFullPath })
        });

        if (response.ok) {
            status.innerText = "Success!";
            // Update UI and Storage
            pushToRecentsStack(targetFullPath); 
        } else {
            status.innerText = "Error: " + response.status;
        }
    } catch (error) {
        console.error(error);
        status.innerText = "Failed";
    }
}
