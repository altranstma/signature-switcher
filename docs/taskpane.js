/* global Office */

const SIGNATURE_MAP_KEY = "signatureMap";

let editingAddress = null; // null = adding new; otherwise the original key being edited

Office.onReady(() => {
    document.getElementById("app").classList.remove("hidden");
    detectCurrentAddress();
    renderSignatureList();
    wireUpUi();
});

function wireUpUi() {
    document.getElementById("useCurrentBtn").addEventListener("click", () => {
        const current = document.getElementById("currentAddress").textContent;
        if (current && current.indexOf("@") > -1) {
            startEdit(current, getSignatureMap()[current.toLowerCase()] || "");
        }
    });

    document.getElementById("saveBtn").addEventListener("click", onSave);
    document.getElementById("cancelBtn").addEventListener("click", resetEditor);

    document.querySelectorAll(".toolbar button").forEach((btn) => {
        btn.addEventListener("click", () => {
            const cmd = btn.getAttribute("data-cmd");
            document.getElementById("signatureEditor").focus();
            if (cmd === "createLink") {
                const url = window.prompt("Link URL:", "https://");
                if (url) {
                    document.execCommand(cmd, false, url);
                }
            } else {
                document.execCommand(cmd, false, null);
            }
        });
    });
}

function detectCurrentAddress() {
    const item = Office.context.mailbox.item;
    if (!item || !item.from) {
        document.getElementById("currentAddress").textContent = "(not available)";
        return;
    }
    item.from.getAsync((result) => {
        const el = document.getElementById("currentAddress");
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.emailAddress) {
            el.textContent = result.value.emailAddress;
        } else {
            el.textContent = "(not available)";
        }
    });
}

function getSignatureMap() {
    const map = Office.context.roamingSettings.get(SIGNATURE_MAP_KEY);
    return map && typeof map === "object" ? map : {};
}

function saveSignatureMap(map, onDone) {
    Office.context.roamingSettings.set(SIGNATURE_MAP_KEY, map);
    Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            showStatus("Could not save: " + result.error.message, true);
        }
        if (onDone) onDone();
    });
}

function renderSignatureList() {
    const map = getSignatureMap();
    const addresses = Object.keys(map).sort();
    const list = document.getElementById("signatureList");
    const emptyState = document.getElementById("emptyState");
    list.innerHTML = "";

    if (addresses.length === 0) {
        emptyState.classList.remove("hidden");
        return;
    }
    emptyState.classList.add("hidden");

    addresses.forEach((address) => {
        const li = document.createElement("li");
        li.className = "signature-item";

        const addrSpan = document.createElement("span");
        addrSpan.className = "signature-item-address";
        addrSpan.textContent = address;
        li.appendChild(addrSpan);

        const actions = document.createElement("span");
        actions.className = "signature-item-actions";

        const editBtn = document.createElement("button");
        editBtn.type = "button";
        editBtn.className = "link-btn";
        editBtn.textContent = "Edit";
        editBtn.addEventListener("click", () => startEdit(address, map[address]));
        actions.appendChild(editBtn);

        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "link-btn danger";
        delBtn.textContent = "Delete";
        delBtn.addEventListener("click", () => onDelete(address));
        actions.appendChild(delBtn);

        li.appendChild(actions);
        list.appendChild(li);
    });
}

function startEdit(address, html) {
    editingAddress = address.toLowerCase();
    document.getElementById("addressInput").value = address;
    document.getElementById("signatureEditor").innerHTML = html || "";
    document.getElementById("editorTitle").textContent = "Edit signature for " + address;
    document.getElementById("cancelBtn").classList.remove("hidden");
    document.getElementById("signatureEditor").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function resetEditor() {
    editingAddress = null;
    document.getElementById("addressInput").value = "";
    document.getElementById("signatureEditor").innerHTML = "";
    document.getElementById("editorTitle").textContent = "Add a signature";
    document.getElementById("cancelBtn").classList.add("hidden");
}

function onSave() {
    const address = document.getElementById("addressInput").value.trim().toLowerCase();
    const html = document.getElementById("signatureEditor").innerHTML.trim();

    if (!address || address.indexOf("@") === -1) {
        showStatus("Enter a valid email address.", true);
        return;
    }
    if (!html) {
        showStatus("Signature can't be empty.", true);
        return;
    }

    const map = getSignatureMap();
    if (editingAddress && editingAddress !== address) {
        delete map[editingAddress];
    }
    map[address] = html;

    saveSignatureMap(map, () => {
        showStatus("Saved signature for " + address + ".", false);
        renderSignatureList();
        resetEditor();
    });
}

function onDelete(address) {
    if (!window.confirm("Delete the signature for " + address + "?")) {
        return;
    }
    const map = getSignatureMap();
    delete map[address.toLowerCase()];
    saveSignatureMap(map, () => {
        showStatus("Deleted signature for " + address + ".", false);
        renderSignatureList();
        if (editingAddress === address.toLowerCase()) {
            resetEditor();
        }
    });
}

function showStatus(message, isError) {
    const el = document.getElementById("statusMessage");
    el.textContent = message;
    el.classList.remove("hidden");
    el.classList.toggle("error", !!isError);
    setTimeout(() => el.classList.add("hidden"), 4000);
}
