async function triggerArchivePipeline(targetFullPath) {
    const standardBtn = document.getElementById("executeBtn");
    const status = document.getElementById("status");
    const item = Office.context.mailbox.item;

    if (!targetFullPath) return;

    // Convert raw EWS ID to Graph-compatible REST ID
    const convertedRestId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
    );

    standardBtn.disabled = true;
    document.querySelectorAll(".quick-btn").forEach(btn => btn.disabled = true);
    
    status.innerText = "Transmitting to engine...";
    status.style.color = "#333";

    try {
        const response = await fetch(EXECUTE_ARCHIVE_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                itemId: convertedRestId, // Map the converted ID here
                targetProject: targetFullPath
            })
        });

        if (response.ok) {
            status.innerText = `Success: Email archived.`;
            status.style.color = "green";
            pushToRecentsStack(targetFullPath);
            renderQuickArchiveUI(); 
        } else {
            throw new Error();
        }
    } catch (error) {
        status.innerText = "Archive failed. Check system connection.";
        status.style.color = "red";
    } finally {
        standardBtn.disabled = (selectedTargetPath === "");
        document.querySelectorAll(".quick-btn").forEach(btn => btn.disabled = false);
    }
}