/* global Office */

const SIGNATURE_MAP_KEY = "signatureMap";
const SOCIAL_ICON_BASE = "https://altranstma.github.io/signature-switcher/assets/social";
const DEFAULT_SOCIAL_ICON_SIZE = 24;
const SOCIAL_PLATFORMS = {
    instagram: { label: "Instagram", icon: `${SOCIAL_ICON_BASE}/instagram.png` },
    facebook: { label: "Facebook", icon: `${SOCIAL_ICON_BASE}/facebook.png` },
    x: { label: "X (Twitter)", icon: `${SOCIAL_ICON_BASE}/x.png` },
    linkedin: { label: "LinkedIn", icon: `${SOCIAL_ICON_BASE}/linkedin.png` },
    youtube: { label: "YouTube", icon: `${SOCIAL_ICON_BASE}/youtube.png` },
    custom: { label: "Custom", icon: null }
};

// Fields the user can toggle on/off in the builder. Each maps to a checkbox
// (id "tg_" + key) and, when on, a directly-editable region in the live
// preview (see buildEditablePreviewHtml/wireEditableFields). "social" is the
// one exception — it reveals the existing repeatable list UI instead of a
// single editable span, since a social link is inherently structured
// (platform + URL), not plain text.
const TOGGLE_FIELDS = ["logo", "name", "titleCompany", "email", "phone", "website", "social", "tagline"];

const FIELD_PLACEHOLDERS = {
    name: "Your Name",
    titleCompany: "Title, Company",
    email: "you@company.com",
    phone: "555-123-4567",
    website: "yourwebsite.com",
    tagline: "Your tagline here"
};

let editingAddress = null; // null = adding new; otherwise the original key being edited
let socialLinks = []; // working list of {id, platform, url, customIconUrl} while editing
let enabledFields = new Set(); // which of TOGGLE_FIELDS are currently checked

// Current text for each editable preview field, kept in sync with the DOM on
// blur (and flushed from the DOM before any rebuild) so a rebuild never loses
// what's currently displayed, even mid-edit.
let builderState = {
    name: "",
    titleCompany: "",
    email: "",
    phone: "",
    website: "",
    tagline: ""
};

Office.onReady(() => {
    document.getElementById("app").classList.remove("hidden");
    wireUpUi();
    resetEditor();
    detectCurrentAddress();
    renderSignatureList();
});

// ---------- Utilities ----------

function escapeHtml(str) {
    return String(str || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function sanitizeHttpUrl(url) {
    const trimmed = String(url || "").trim();
    if (/^https?:\/\//i.test(trimmed)) {
        return escapeHtml(trimmed);
    }
    return "";
}

function sanitizePhoneForTel(phone) {
    return String(phone || "").replace(/[^\d+]/g, "");
}

// ---------- Signature HTML builder (canonical output used for saving) ----------

function renderSocialIconsHtml(links, iconSize) {
    return (links || [])
        .filter(s => s.url)
        .map(s => {
            const icon = s.platform === "custom" ? sanitizeHttpUrl(s.customIconUrl) : (SOCIAL_PLATFORMS[s.platform] || {}).icon;
            const linkUrl = sanitizeHttpUrl(s.url);
            if (!icon || !linkUrl) return "";
            return `<a href="${linkUrl}" style="text-decoration:none;margin-right:10px;"><img src="${icon}" width="${iconSize}" height="${iconSize}" style="border:0;vertical-align:middle;" alt="" /></a>`;
        })
        .join("");
}

function buildSignatureHtml(fields) {
    const logoCell = fields.logoUrl
        ? `<td style="vertical-align:middle;padding-right:14px;"><img src="${sanitizeHttpUrl(fields.logoUrl)}" width="70" style="display:block;border:0;" alt="" /></td>`
        : "";

    const contactParts = [];
    if (fields.email) {
        contactParts.push(`<a href="mailto:${escapeHtml(fields.email)}" style="color:#1155cc;text-decoration:underline;">${escapeHtml(fields.email)}</a>`);
    }
    if (fields.phone) {
        contactParts.push(`<a href="tel:${sanitizePhoneForTel(fields.phone)}" style="color:#333333;text-decoration:none;">${escapeHtml(fields.phone)}</a>`);
    }
    if (fields.website) {
        const url = sanitizeHttpUrl(fields.website);
        contactParts.push(url
            ? `<a href="${url}" style="color:#1155cc;text-decoration:underline;">${escapeHtml(fields.website.replace(/^https?:\/\//i, ""))}</a>`
            : escapeHtml(fields.website));
    }
    const contactLine = contactParts.join(' <span style="color:#999999;">|</span> ');

    const titleCompany = [fields.title, fields.company].filter(Boolean).join(", ");

    const iconSize = fields.socialIconSize || DEFAULT_SOCIAL_ICON_SIZE;
    const socialImgs = renderSocialIconsHtml(fields.socialLinks, iconSize);

    let html = `<table cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif;border-collapse:collapse;"><tr>`;
    html += logoCell;
    html += `<td style="${fields.logoUrl ? "border-left:2px solid #cccccc;padding-left:14px;" : ""}vertical-align:middle;">`;
    if (fields.name) html += `<div style="font-size:12pt;font-weight:bold;color:#222222;">${escapeHtml(fields.name)}</div>`;
    if (titleCompany) html += `<div style="font-size:11pt;font-weight:bold;color:#666666;">${escapeHtml(titleCompany)}</div>`;
    if (contactLine) html += `<div style="font-size:11pt;color:#333333;margin-top:3px;">${contactLine}</div>`;
    if (socialImgs) html += `<div style="margin-top:8px;">${socialImgs}</div>`;
    html += `</td></tr>`;
    if (fields.tagline && fields.tagline.text) {
        const taglineUrl = sanitizeHttpUrl(fields.tagline.url);
        const taglineContent = taglineUrl
            ? `<a href="${taglineUrl}" style="color:#1155cc;text-decoration:underline;">${escapeHtml(fields.tagline.text)}</a>`
            : escapeHtml(fields.tagline.text);
        html += `<tr><td colspan="2" style="padding-top:8px;font-size:11pt;color:#555555;">${taglineContent}</td></tr>`;
    }
    html += `</table>`;
    return html;
}

// ---------- Builder state <-> saved fields shape ----------

function collectBuilderFields() {
    flushPreviewToState();
    return {
        logoUrl: enabledFields.has("logo") ? document.getElementById("f_logoUrl").value.trim() : "",
        name: enabledFields.has("name") ? builderState.name : "",
        title: enabledFields.has("titleCompany") ? builderState.titleCompany : "",
        company: "",
        email: enabledFields.has("email") ? builderState.email : "",
        phone: enabledFields.has("phone") ? builderState.phone : "",
        website: enabledFields.has("website") ? builderState.website : "",
        socialLinks: enabledFields.has("social") ? collectSocialLinks() : [],
        socialIconSize: parseInt(document.getElementById("f_socialIconSize").value, 10) || DEFAULT_SOCIAL_ICON_SIZE,
        tagline: {
            text: enabledFields.has("tagline") ? builderState.tagline : "",
            url: enabledFields.has("tagline") ? document.getElementById("f_taglineUrl").value.trim() : ""
        }
    };
}

function populateBuilderFields(fields) {
    fields = fields || {};
    builderState = {
        name: fields.name || "",
        titleCompany: fields.titleCompany || [fields.title, fields.company].filter(Boolean).join(", "),
        email: fields.email || "",
        phone: fields.phone || "",
        website: fields.website || "",
        tagline: (fields.tagline && fields.tagline.text) || ""
    };

    document.getElementById("f_logoUrl").value = fields.logoUrl || "";
    document.getElementById("f_socialIconSize").value = fields.socialIconSize || DEFAULT_SOCIAL_ICON_SIZE;
    document.getElementById("f_taglineUrl").value = (fields.tagline || {}).url || "";

    enabledFields = new Set();
    if (fields.logoUrl) enabledFields.add("logo");
    if (builderState.name) enabledFields.add("name");
    if (builderState.titleCompany) enabledFields.add("titleCompany");
    if (builderState.email) enabledFields.add("email");
    if (builderState.phone) enabledFields.add("phone");
    if (builderState.website) enabledFields.add("website");
    if (fields.socialLinks && fields.socialLinks.length) enabledFields.add("social");
    if (builderState.tagline) enabledFields.add("tagline");

    TOGGLE_FIELDS.forEach(key => {
        document.getElementById("tg_" + key).checked = enabledFields.has(key);
    });
    updateConditionalRowsVisibility();

    document.getElementById("socialLinksList").innerHTML = "";
    socialLinks = [];
    if (fields.socialLinks && fields.socialLinks.length) {
        fields.socialLinks.forEach(s => addSocialLinkRow(s));
    } else {
        addSocialLinkRow();
    }

    renderPreview();
}

function clearBuilderFields() {
    populateBuilderFields({});
}

function updateConditionalRowsVisibility() {
    document.getElementById("logoUrlRow").classList.toggle("hidden", !enabledFields.has("logo"));
    document.getElementById("socialLinksSection").classList.toggle("hidden", !enabledFields.has("social"));
    document.getElementById("taglineUrlRow").classList.toggle("hidden", !enabledFields.has("tagline"));
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

    // Order is whatever order the rows sit in the DOM — collectSocialLinks()
    // reads them in that order, so moving a row is enough to reorder; no
    // separate index needs to be tracked or persisted.
    const moveUpBtn = document.createElement("button");
    moveUpBtn.type = "button";
    moveUpBtn.className = "link-btn social-move";
    moveUpBtn.textContent = "▲";
    moveUpBtn.title = "Move up";

    const moveDownBtn = document.createElement("button");
    moveDownBtn.type = "button";
    moveDownBtn.className = "link-btn social-move";
    moveDownBtn.textContent = "▼";
    moveDownBtn.title = "Move down";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "link-btn danger social-remove";
    removeBtn.textContent = "Remove";

    select.addEventListener("change", () => {
        customIconInput.style.display = select.value === "custom" ? "" : "none";
        renderPreview();
    });
    urlInput.addEventListener("input", renderPreview);
    customIconInput.addEventListener("input", renderPreview);
    moveUpBtn.addEventListener("click", () => {
        const prev = row.previousElementSibling;
        if (prev) {
            row.parentNode.insertBefore(row, prev);
            renderPreview();
        }
    });
    moveDownBtn.addEventListener("click", () => {
        const next = row.nextElementSibling;
        if (next) {
            row.parentNode.insertBefore(next, row);
            renderPreview();
        }
    });
    removeBtn.addEventListener("click", () => {
        row.remove();
        socialLinks = socialLinks.filter(s => s.id !== id);
        renderPreview();
    });

    row.appendChild(select);
    row.appendChild(urlInput);
    row.appendChild(customIconInput);
    row.appendChild(moveUpBtn);
    row.appendChild(moveDownBtn);
    row.appendChild(removeBtn);
    document.getElementById("socialLinksList").appendChild(row);
    renderPreview();
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

// ---------- Editable live preview ----------

// Copies whatever's currently displayed in each editable preview field back
// into builderState. Must run before any preview rebuild (renderPreview) and
// before reading fields for save, so an in-progress edit that hasn't been
// blurred yet is never lost.
function flushPreviewToState() {
    document.querySelectorAll("#builderPreview [data-field]").forEach((el) => {
        const key = el.dataset.field;
        if (!(key in builderState)) return;
        builderState[key] = el.classList.contains("placeholder-text") ? "" : el.textContent.trim();
    });
}

function fieldSpan(key) {
    if (!enabledFields.has(key)) return "";
    const placeholder = FIELD_PLACEHOLDERS[key];
    const value = builderState[key];
    const isPlaceholder = !value;
    const display = escapeHtml(isPlaceholder ? placeholder : value);
    const cls = isPlaceholder ? "preview-editable placeholder-text" : "preview-editable";
    return `<span class="${cls}" contenteditable="true" data-field="${key}" data-placeholder="${escapeHtml(placeholder)}">${display}</span>`;
}

function buildEditablePreviewHtml() {
    const hasLogo = enabledFields.has("logo");
    const logoUrl = document.getElementById("f_logoUrl").value.trim();
    let logoCell = "";
    if (hasLogo) {
        logoCell = logoUrl
            ? `<td style="vertical-align:middle;padding-right:14px;"><img src="${sanitizeHttpUrl(logoUrl)}" width="70" style="display:block;border:0;" alt="" /></td>`
            : `<td style="vertical-align:middle;padding-right:14px;"><span class="preview-hint">Add a logo URL below &darr;</span></td>`;
    }

    const contactPieces = [];
    if (enabledFields.has("email")) contactPieces.push(fieldSpan("email"));
    if (enabledFields.has("phone")) contactPieces.push(fieldSpan("phone"));
    if (enabledFields.has("website")) contactPieces.push(fieldSpan("website"));
    const contactLine = contactPieces.join(' <span style="color:#999999;">|</span> ');

    const iconSize = parseInt(document.getElementById("f_socialIconSize").value, 10) || DEFAULT_SOCIAL_ICON_SIZE;
    let socialHtml = "";
    if (enabledFields.has("social")) {
        const links = collectSocialLinks();
        socialHtml = links.length
            ? renderSocialIconsHtml(links, iconSize)
            : `<span class="preview-hint">Add a social link below &darr;</span>`;
    }

    let html = `<table cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif;border-collapse:collapse;"><tr>`;
    html += logoCell;
    html += `<td style="${hasLogo ? "border-left:2px solid #cccccc;padding-left:14px;" : ""}vertical-align:middle;">`;
    if (enabledFields.has("name")) {
        html += `<div style="font-size:12pt;font-weight:bold;color:#222222;">${fieldSpan("name")}</div>`;
    }
    if (enabledFields.has("titleCompany")) {
        html += `<div style="font-size:11pt;font-weight:bold;color:#666666;">${fieldSpan("titleCompany")}</div>`;
    }
    if (contactLine) {
        html += `<div style="font-size:11pt;color:#333333;margin-top:3px;">${contactLine}</div>`;
    }
    if (socialHtml) {
        html += `<div style="margin-top:8px;">${socialHtml}</div>`;
    }
    if (enabledFields.size === 0) {
        html += `<div class="preview-hint">Check a box above to start building your signature.</div>`;
    }
    html += `</td></tr>`;
    if (enabledFields.has("tagline")) {
        html += `<tr><td colspan="2" style="padding-top:8px;font-size:11pt;color:#555555;">${fieldSpan("tagline")}</td></tr>`;
    }
    html += `</table>`;
    return html;
}

function wireEditableFields() {
    document.querySelectorAll("#builderPreview [contenteditable]").forEach((el) => {
        el.addEventListener("focus", () => {
            if (el.classList.contains("placeholder-text")) {
                el.textContent = "";
                el.classList.remove("placeholder-text");
            }
        });
        el.addEventListener("blur", () => {
            const key = el.dataset.field;
            const value = el.textContent.trim();
            builderState[key] = value;
            if (!value) {
                el.textContent = el.dataset.placeholder;
                el.classList.add("placeholder-text");
            }
        });
    });
}

function renderPreview() {
    flushPreviewToState();
    const preview = document.getElementById("builderPreview");
    preview.innerHTML = buildEditablePreviewHtml();
    wireEditableFields();
}

// ---------- Wire up ----------

function wireUpUi() {
    document.getElementById("builderTab").addEventListener("click", () => setMode("builder"));
    document.getElementById("customTab").addEventListener("click", () => setMode("custom"));
    document.getElementById("addSocialLinkBtn").addEventListener("click", () => addSocialLinkRow());

    TOGGLE_FIELDS.forEach(key => {
        document.getElementById("tg_" + key).addEventListener("change", (e) => {
            if (e.target.checked) {
                enabledFields.add(key);
            } else {
                enabledFields.delete(key);
            }
            updateConditionalRowsVisibility();
            renderPreview();
        });
    });

    ["f_logoUrl", "f_socialIconSize", "f_taglineUrl"]
        .forEach(id => document.getElementById(id).addEventListener("input", renderPreview));

    document.getElementById("useCurrentBtn").addEventListener("click", () => {
        const current = document.getElementById("currentAddress").textContent;
        if (current && current.indexOf("@") > -1) {
            loadAddressIntoEditor(current);
        }
    });

    document.getElementById("saveBtn").addEventListener("click", onSave);
    document.getElementById("cancelBtn").addEventListener("click", resetEditor);

    document.getElementById("editAsCustomBtn").addEventListener("click", () => {
        document.getElementById("signatureEditor").innerHTML = buildSignatureHtml(collectBuilderFields());
        setMode("custom");
    });

    let savedSelectionRange = null;

    document.querySelectorAll(".toolbar button").forEach((btn) => {
        btn.addEventListener("click", () => {
            const cmd = btn.getAttribute("data-cmd");
            document.getElementById("signatureEditor").focus();
            if (cmd === "createLink") {
                // window.prompt() is silently suppressed in some Office Add-in
                // task pane hosts, so use an inline input instead of a native
                // dialog. Save the selection first since focusing the input
                // below would otherwise collapse it.
                const selection = window.getSelection();
                savedSelectionRange = selection.rangeCount ? selection.getRangeAt(0) : null;
                document.getElementById("linkUrlInput").value = "https://";
                document.getElementById("linkUrlRow").classList.remove("hidden");
                document.getElementById("linkUrlInput").focus();
                document.getElementById("linkUrlInput").select();
            } else {
                document.execCommand(cmd, false, null);
            }
        });
    });

    function closeLinkUrlRow() {
        document.getElementById("linkUrlRow").classList.add("hidden");
        savedSelectionRange = null;
    }

    document.getElementById("linkUrlInsertBtn").addEventListener("click", () => {
        const url = document.getElementById("linkUrlInput").value.trim();
        if (url) {
            const editor = document.getElementById("signatureEditor");
            editor.focus();
            const selection = window.getSelection();
            if (savedSelectionRange) {
                selection.removeAllRanges();
                selection.addRange(savedSelectionRange);
            }
            document.execCommand("createLink", false, url);
        }
        closeLinkUrlRow();
    });
    document.getElementById("linkUrlCancelBtn").addEventListener("click", closeLinkUrlRow);
    document.getElementById("linkUrlInput").addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            document.getElementById("linkUrlInsertBtn").click();
        } else if (e.key === "Escape") {
            closeLinkUrlRow();
        }
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
        editBtn.addEventListener("click", () => loadAddressIntoEditor(address));
        actions.appendChild(editBtn);

        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "link-btn danger";
        delBtn.textContent = "Delete";
        delBtn.addEventListener("click", () => onDelete(address, delBtn));
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
            showStatus("At least check and enter a name for the builder signature.", true);
            return;
        }
        entry = { mode: "builder", fields, html: buildSignatureHtml(fields) };
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

// window.confirm() is silently suppressed in some Office Add-in task pane
// hosts (notably OWA's), so a native confirm dialog never appears and the
// delete just looks broken. Require a second click within a few seconds
// instead of a native dialog.
function onDelete(address, btnEl) {
    if (btnEl.dataset.confirming !== "1") {
        btnEl.dataset.confirming = "1";
        const originalLabel = btnEl.textContent;
        btnEl.textContent = "Confirm delete?";
        btnEl.dataset.resetTimer = setTimeout(() => {
            btnEl.dataset.confirming = "0";
            btnEl.textContent = originalLabel;
        }, 4000);
        return;
    }

    clearTimeout(Number(btnEl.dataset.resetTimer));
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
