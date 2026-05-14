// Configuration endpoints
const GET_PROJECTS_URL = "https://default062a8e8e449048f39ee3b309e2cfa4.ad.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/405ddbd55c224c9ebe1d2bc5b85a6597/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=P9Tu-84M5_ZRI2lryh6GQTPq9erJ9yTd9JNk0CVZli4";
const EXECUTE_ARCHIVE_URL = "https://default062a8e8e449048f39ee3b309e2cfa4.ad.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/36e5dea2ad0f4486ac1c61e45e6dde4d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vSvhWWzURBFmq2LDhc7ysp6wZ9blVACL2UNQ2SNRarA";

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadProjects();
        document.getElementById("executeBtn").onclick = archiveEmail;
    }
});

async function loadProjects() {
    const select = document.getElementById("projectSelect");
    const button = document.getElementById("executeBtn");
    const status = document.getElementById("status");

    try {
        const response = await fetch(GET_PROJECTS_URL);
        if (!response.ok) throw new Error("Failed to fetch folder tree.");
        
        const projects = await response.json();
        
        select.innerHTML = "";
        projects.forEach(project => {
            const opt = document.createElement('option');
            opt.value = project;
            opt.innerHTML = project;
            select.appendChild(opt);
        });
        
        button.disabled = false;
    } catch (error) {
        status.innerText = "Error loading project directories.";
        status.style.color = "red";
        select.innerHTML = '<option value="">Load failed</option>';
    }
}

async function archiveEmail() {
    const button = document.getElementById("executeBtn");
    const status = document.getElementById("status");
    const project = document.getElementById("projectSelect").value;
    const item = Office.context.mailbox.item;

    if (!project) {
        status.innerText = "Select a valid project.";
        return;
    }

    button.disabled = true;
    status.innerText = "Transmitting to engine...";
    status.style.color = "black";

    const payload = {
        itemId: item.itemId, // Standard EWS ID used by Power Automate Outlook connector
        targetProject: project
    };

    try {
        const response = await fetch(EXECUTE_ARCHIVE_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            status.innerText = `Archived successfully to /${project}/`;
            status.style.color = "green";
        } else {
            throw new Error("Pipeline rejected payload.");
        }
    } catch (error) {
        status.innerText = "Execution failed. Check backend flow history.";
        status.style.color = "red";
        button.disabled = false;
    }
}