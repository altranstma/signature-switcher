/* global Office, SOCIAL_PLATFORMS, buildSignatureHtml */

const SIGNATURE_MAP_KEY = "signatureMap";
const LOG_PREFIX = "[SigSwitcher-Taskpane]";

function byteSize(value) {
    return new Blob([JSON.stringify(value)]).size;
}

let editingAddress = null; // null = adding new; otherwise the original key being edited
let socialLinks = []; // working list of {id, platform, url, customIconUrl} while editing

Office.onReady(() => {
    document.getElementById("app").classList.remove("hidden");
    detectCurrentAddress();
    renderSignatureList();
    wireUpUi();
    addSocialLinkRow(); // start with one empty row
    updateBuilderPreview();
});

// ---------- Utilities ----------
// escapeHtml, sanitizeHttpUrl, sanitizePhoneForTel, and buildSignatureHtml
// live in signature-html.js, shared with launchevent.js.

function collectBuilderFields() {
    return {
        logoUrl: document.getElementById("f_logoUrl").value.trim(),
        name: document.getElementById("f_name").value.trim(),
        title: document.getElementById("f_title").value.trim(),
        company: document.getElementById("f_company").value.trim(),
        email: document.getElementById("f_email").value.trim(),
        phone: document.getElementById("f_phone").value.trim(),
        website: document.getElementById("f_website").value.trim(),
        socialLinks: collectSocialLinks(),
        tagline: {
            text: document.getElementById("f_taglineText").value.trim(),
            url: document.getElementById("f_taglineUrl").value.trim()
        }
    };
}

function populateBuilderFields(fields) {
    document.getElementById("f_logoUrl").value = fields.logoUrl || "";
    document.getElementById("f_name").value = fields.name || "";
    document.getElementById("f_title").value = fields.title || "";
    document.getElementById("f_company").value = fields.company || "";
    document.getElementById("f_email").value = fields.email || "";
    document.getElementById("f_phone").value = fields.phone || "";
    document.getElementById("f_website").value = fields.website || "";
    document.getElementById("f_taglineText").value = (fields.tagline || {}).text || "";
    document.getElementById("f_taglineUrl").value = (fields.tagline || {}).url || "";

    document.getElementById("socialLinksList").innerHTML = "";
    socialLinks = [];
    if (fields.socialLinks && fields.socialLinks.length) {
        fields.socialLinks.forEach(s => addSocialLinkRow(s));
    } else {
        addSocialLinkRow();
    }
    updateBuilderPreview();
}

function clearBuilderFields() {
    populateBuilderFields({});
}

// ---------- Social links repeatable rows ----------

function addSocialLinkRow(existing) {
    const id = "social_" + Math.floor(Math.random() * 1e9) + "_" + socialLinks.length;
    const entry = existing ? { ...existing } : { platform: "instagram", url: "", customIconUrl: "" };
    socialLinks.push({ id, ...entry });

    const row = document.createElement("div");
    row.className = "social-row";
    row.dataset.id = id;

    const select = document.createElement("select");
    select.className = "social-platform";
    Object.keys(SOCIAL_PLATFORMS).forEach(key => {
        const opt = document.createElement("option");
        opt.value = key;
        opt.textContent = SOCIAL_PLATFORMS[key].label;
        if (key === entry.platform) opt.selected = true;
        select.appendChild(opt);
    });

    const urlInput = document.createElement("input");
    urlInput.type = "text";
    urlInput.className = "social-url";
    urlInput.placeholder = "Profile URL";
    urlInput.value = entry.url || "";

    const customIconInput = document.createElement("input");
    customIconInput.type = "text";
    customIconInput.className = "social-custom-icon";
    customIconInput.placeholder = "Custom icon image URL";
    customIconInput.value = entry.customIconUrl || "";
    customIconInput.style.display = entry.platform === "custom" ? "" : "none";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "link-btn danger social-remove";
    removeBtn.textContent = "Remove";

    select.addEventListener("change", () => {
        customIconInput.style.display = select.value === "custom" ? "" : "none";
        updateBuilderPreview();
    });
    urlInput.addEventListener("input", updateBuilderPreview);
    customIconInput.addEventListener("input", updateBuilderPreview);
    removeBtn.addEventListener("click", () => {
        row.remove();
        socialLinks = socialLinks.filter(s => s.id !== id);
        updateBuilderPreview();
    });

    row.appendChild(select);
    row.appendChild(urlInput);
    row.appendChild(customIconInput);
    row.appendChild(removeBtn);
    document.getElementById("socialLinksList").appendChild(row);
    updateBuilderPreview();
}

function collectSocialLinks() {
    const rows = document.querySelectorAll("#socialLinksList .social-row");
    const result = [];
    rows.forEach(row => {
        const platform = row.querySelector(".social-platform").value;
        const url = row.querySelector(".social-url").value.trim();
        const customIconUrl = row.querySelector(".social-custom-icon").value.trim();
        if (url) result.push({ platform, url, customIconUrl });
    });
    return result;
}

// ---------- Mode switching ----------

function setMode(mode) {
    document.getElementById("builderTab").classList.toggle("active", mode === "builder");
    document.getElementById("customTab").classList.toggle("active", mode === "custom");
    document.getElementById("builderPane").classList.toggle("hidden", mode !== "builder");
    document.getElementById("customPane").classList.toggle("hidden", mode !== "custom");
}

function currentMode() {
    return document.getElementById("builderPane").classList.contains("hidden") ? "custom" : "builder";
}

function updateBuilderPreview() {
    const fields = collectBuilderFields();
    document.getElementById("builderPreview").innerHTML = buildSignatureHtml(fields);
}

// ---------- Wire up ----------

function wireUpUi() {
    document.getElementById("builderTab").addEventListener("click", () => setMode("builder"));
    document.getElementById("customTab").addEventListener("click", () => setMode("custom"));
    document.getElementById("addSocialLinkBtn").addEventListener("click", () => addSocialLinkRow());

    ["f_logoUrl", "f_name", "f_title", "f_company", "f_email", "f_phone", "f_website", "f_taglineText", "f_taglineUrl"]
        .forEach(id => document.getElementById(id).addEventListener("input", updateBuilderPreview));

    document.getElementById("useCurrentBtn").addEventListener("click", () => {
        const current = document.getElementById("currentAddress").textContent;
        if (current && current.indexOf("@") > -1) {
            loadAddressIntoEditor(current);
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
                if (url) document.execCommand(cmd, false, url);
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

// Office.context.roamingSettings has a hard 32KB-per-user cap. Older builds of
// this add-in stored a redundant rendered `html` copy alongside `fields` for
// every builder-mode entry, which burned through it after only ~10 addresses.
// Strip those leftover copies from the whole map on every save so accounts
// that already hit the cap reclaim space immediately, not just on next edit.
function compactSignatureMap(map) {
    Object.keys(map).forEach((address) => {
        const entry = map[address];
        if (entry && typeof entry === "object" && entry.mode !== "custom" && entry.fields) {
            delete entry.html;
        }
    });
    return map;
}

function saveSignatureMap(map, onDone) {
    const beforeBytes = byteSize(map);
    compactSignatureMap(map);
    const afterBytes = byteSize(map);
    const perEntryBytes = Object.keys(map)
        .map((address) => ({ address, bytes: byteSize(map[address]) }))
        .sort((a, b) => b.bytes - a.bytes);
    console.log(LOG_PREFIX, "map bytes before/after compaction:", beforeBytes, "/", afterBytes);
    console.log(LOG_PREFIX, "per-entry bytes (largest first):", JSON.stringify(perEntryBytes));

    Office.context.roamingSettings.set(SIGNATURE_MAP_KEY, map);
    Office.context.roamingSettings.saveAsync((result) => {
        console.log(LOG_PREFIX, "roamingSettings.saveAsync ->", result.status, result.error ? JSON.stringify(result.error) : "");
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
        editBtn.addEventListener("click", () => loadAddressIntoEditor(address));
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

function loadAddressIntoEditor(address) {
    const map = getSignatureMap();
    const entry = map[address.toLowerCase()];
    editingAddress = address.toLowerCase();
    document.getElementById("addressInput").value = address;
    document.getElementById("editorTitle").textContent = "Edit signature for " + address;
    document.getElementById("cancelBtn").classList.remove("hidden");

    if (entry && entry.mode === "custom") {
        setMode("custom");
        document.getElementById("signatureEditor").innerHTML = entry.html || "";
    } else if (entry && entry.fields) {
        setMode("builder");
        populateBuilderFields(entry.fields);
    } else {
        setMode("builder");
        clearBuilderFields();
    }
    document.getElementById("editorSection").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function resetEditor() {
    editingAddress = null;
    document.getElementById("addressInput").value = "";
    document.getElementById("editorTitle").textContent = "Add a signature";
    document.getElementById("cancelBtn").classList.add("hidden");
    document.getElementById("signatureEditor").innerHTML = "";
    clearBuilderFields();
    setMode("builder");
}

function onSave() {
    const address = document.getElementById("addressInput").value.trim().toLowerCase();
    if (!address || address.indexOf("@") === -1) {
        showStatus("Enter a valid email address.", true);
        return;
    }

    const mode = currentMode();
    let entry;
    if (mode === "custom") {
        const html = document.getElementById("signatureEditor").innerHTML.trim();
        if (!html) {
            showStatus("Signature can't be empty.", true);
            return;
        }
        entry = { mode: "custom", html };
    } else {
        const fields = collectBuilderFields();
        if (!fields.name) {
            showStatus("At least enter a name for the builder signature.", true);
            return;
        }
        entry = { mode: "builder", fields };
    }

    const map = getSignatureMap();
    if (editingAddress && editingAddress !== address) {
        delete map[editingAddress];
    }
    map[address] = entry;

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
