/* global Office */

const SIGNATURE_MAP_KEY = "signatureMap";
const SOCIAL_ICON_BASE = "https://altranstma.github.io/signature-switcher/assets/social";
const DEFAULT_SOCIAL_ICON_SIZE = 24;
const CLIENT_LOGO_BASE = "https://altrans.net/images/clientlogo/";

// urlTemplate turns a plain username/handle into a full profile URL, so the
// user only has to type e.g. "danieloliver" instead of the whole link.
// "custom" has no template — that option is inherently a full URL + custom
// icon, since there's no known base URL to build from.
const SOCIAL_PLATFORMS = {
    instagram: { label: "Instagram", icon: `${SOCIAL_ICON_BASE}/instagram.png`, urlTemplate: "https://instagram.com/{u}" },
    facebook: { label: "Facebook", icon: `${SOCIAL_ICON_BASE}/facebook.png`, urlTemplate: "https://facebook.com/{u}" },
    x: { label: "X (Twitter)", icon: `${SOCIAL_ICON_BASE}/x.png`, urlTemplate: "https://x.com/{u}" },
    linkedin: { label: "LinkedIn", icon: `${SOCIAL_ICON_BASE}/linkedin.png`, urlTemplate: "https://linkedin.com/in/{u}" },
    youtube: { label: "YouTube", icon: `${SOCIAL_ICON_BASE}/youtube.png`, urlTemplate: "https://youtube.com/@{u}" },
    custom: { label: "Custom", icon: null, urlTemplate: null }
};

let editingAddress = null; // null = adding new; otherwise the original key being edited
let socialLinks = []; // working list of {id, platform, url, customIconUrl} while editing
let scalePercent = 100; // "Overall size" stepper, 50-200
let spacingPercent = 100; // "Line spacing" stepper, 50-200

Office.onReady(() => {
    document.getElementById("app").classList.remove("hidden");
    populateLogoPicker();
    detectCurrentAddress();
    renderSignatureList();
    wireUpUi();
    addSocialLinkRow(); // start with one empty row
    updateBuilderPreview();
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

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

// Turns a social row's stored value into a full profile URL. A value that's
// already a full URL (either the "custom" platform, or an older saved
// signature from before usernames-only) is used as-is; otherwise it's
// treated as a bare username/handle and expanded via the platform's
// urlTemplate.
function resolveSocialUrl(platform, value) {
    const trimmed = String(value || "").trim();
    if (!trimmed) return "";
    if (/^https?:\/\//i.test(trimmed)) {
        return trimmed;
    }
    const info = SOCIAL_PLATFORMS[platform];
    if (!info || !info.urlTemplate) return "";
    const handle = trimmed.replace(/^[@/]+/, "");
    return info.urlTemplate.replace("{u}", encodeURIComponent(handle));
}

// ---------- Signature HTML builder ----------

function buildSignatureHtml(fields) {
    const scale = fields.scale || 1;
    const spacing = fields.lineSpacing || 1;
    const pt = (base) => {
        const v = Math.round(base * scale * 10) / 10;
        return (v % 1 === 0) ? String(v) : v.toFixed(1);
    };
    const px = (base) => Math.round(base * scale);
    const sp = (base) => Math.round(base * spacing);

    const logoCell = fields.logoUrl
        ? `<td style="vertical-align:middle;padding-right:14px;"><img src="${sanitizeHttpUrl(fields.logoUrl)}" width="${px(70)}" style="display:block;border:0;" alt="" /></td>`
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

    const titleCompanyParts = [fields.title, fields.company].filter(Boolean).map(escapeHtml);
    const titleCompany = titleCompanyParts.join(' <span style="color:#999999;">|</span> ');

    const iconSize = px(fields.socialIconSize || DEFAULT_SOCIAL_ICON_SIZE);
    const socialImgs = (fields.socialLinks || [])
        .filter(s => s.url)
        .map(s => {
            const icon = s.platform === "custom" ? sanitizeHttpUrl(s.customIconUrl) : (SOCIAL_PLATFORMS[s.platform] || {}).icon;
            const linkUrl = sanitizeHttpUrl(resolveSocialUrl(s.platform, s.url));
            if (!icon || !linkUrl) return "";
            return `<a href="${linkUrl}" style="text-decoration:none;margin-right:10px;"><img src="${icon}" width="${iconSize}" height="${iconSize}" style="border:0;vertical-align:middle;" alt="" /></a>`;
        })
        .join("");

    let html = `<table cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif;border-collapse:collapse;"><tr>`;
    html += logoCell;
    html += `<td style="${fields.logoUrl ? "border-left:2px solid #cccccc;padding-left:14px;" : ""}vertical-align:middle;">`;
    if (fields.name) html += `<div style="font-size:${pt(12)}pt;font-weight:bold;color:#222222;">${escapeHtml(fields.name)}</div>`;
    if (titleCompany) html += `<div style="font-size:${pt(10)}pt;font-weight:bold;color:#666666;margin-top:${sp(2)}px;">${titleCompany}</div>`;
    if (contactLine) html += `<div style="font-size:${pt(8.5)}pt;color:#333333;margin-top:${sp(3)}px;">${contactLine}</div>`;
    if (socialImgs) html += `<div style="margin-top:${sp(8)}px;">${socialImgs}</div>`;
    html += `</td></tr>`;
    if (fields.tagline && fields.tagline.text) {
        const taglineUrl = sanitizeHttpUrl(fields.tagline.url);
        const taglineContent = taglineUrl
            ? `<a href="${taglineUrl}" style="color:#1155cc;text-decoration:underline;">${escapeHtml(fields.tagline.text)}</a>`
            : escapeHtml(fields.tagline.text);
        html += `<tr><td colspan="2" style="padding-top:${sp(8)}px;font-size:${pt(8.5)}pt;color:#555555;">${taglineContent}</td></tr>`;
    }
    html += `</table>`;
    return html;
}

function collectBuilderFields() {
    return {
        logoUrl: document.getElementById("tg_logo").checked ? document.getElementById("f_logoUrl").value.trim() : "",
        name: document.getElementById("f_name").value.trim(),
        title: document.getElementById("f_title").value.trim(),
        company: document.getElementById("f_company").value.trim(),
        email: document.getElementById("tg_email").checked ? document.getElementById("f_email").value.trim() : "",
        phone: document.getElementById("tg_phone").checked ? document.getElementById("f_phone").value.trim() : "",
        website: document.getElementById("tg_website").checked ? document.getElementById("f_website").value.trim() : "",
        socialLinks: document.getElementById("tg_social").checked ? collectSocialLinks() : [],
        socialIconSize: DEFAULT_SOCIAL_ICON_SIZE,
        tagline: {
            text: document.getElementById("tg_tagline").checked ? document.getElementById("f_taglineText").value.trim() : "",
            url: document.getElementById("tg_tagline").checked ? document.getElementById("f_taglineUrl").value.trim() : ""
        },
        scale: scalePercent / 100,
        lineSpacing: spacingPercent / 100
    };
}

function populateBuilderFields(fields) {
    fields = fields || {};
    document.getElementById("f_name").value = fields.name || "";
    document.getElementById("f_title").value = fields.title || "";
    document.getElementById("f_company").value = fields.company || "";
    document.getElementById("f_email").value = fields.email || "";
    document.getElementById("f_phone").value = fields.phone || "";
    document.getElementById("f_website").value = fields.website || "";
    document.getElementById("f_taglineText").value = (fields.tagline || {}).text || "";
    document.getElementById("f_taglineUrl").value = (fields.tagline || {}).url || "";

    setLogoUrl(fields.logoUrl || "");

    document.getElementById("tg_logo").checked = !!fields.logoUrl;
    document.getElementById("tg_email").checked = !!fields.email;
    document.getElementById("tg_phone").checked = !!fields.phone;
    document.getElementById("tg_website").checked = !!fields.website;
    document.getElementById("tg_social").checked = !!(fields.socialLinks && fields.socialLinks.length);
    document.getElementById("tg_tagline").checked = !!((fields.tagline || {}).text);
    updateConditionalVisibility();

    scalePercent = clamp(Math.round((fields.scale || 1) * 100), 50, 200);
    spacingPercent = clamp(Math.round((fields.lineSpacing || 1) * 100), 50, 200);
    document.getElementById("scaleValue").textContent = scalePercent + "%";
    document.getElementById("spacingValue").textContent = spacingPercent + "%";

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

function updateConditionalVisibility() {
    document.getElementById("logoSection").classList.toggle("hidden", !document.getElementById("tg_logo").checked);
    document.getElementById("emailSection").classList.toggle("hidden", !document.getElementById("tg_email").checked);
    document.getElementById("phoneSection").classList.toggle("hidden", !document.getElementById("tg_phone").checked);
    document.getElementById("websiteSection").classList.toggle("hidden", !document.getElementById("tg_website").checked);
    document.getElementById("socialSection").classList.toggle("hidden", !document.getElementById("tg_social").checked);
    document.getElementById("taglineSection").classList.toggle("hidden", !document.getElementById("tg_tagline").checked);
}

// ---------- Logo picker ----------

function populateLogoPicker() {
    const picker = document.getElementById("f_logoPicker");
    picker.innerHTML = "";

    const placeholderOpt = document.createElement("option");
    placeholderOpt.value = "";
    placeholderOpt.textContent = "— Select a logo —";
    picker.appendChild(placeholderOpt);

    (window.CLIENT_LOGOS || []).forEach(logo => {
        const opt = document.createElement("option");
        opt.value = logo.url || (CLIENT_LOGO_BASE + logo.file);
        opt.textContent = logo.name || logo.file;
        picker.appendChild(opt);
    });

    const customOpt = document.createElement("option");
    customOpt.value = "__custom__";
    customOpt.textContent = "Custom URL…";
    picker.appendChild(customOpt);

    picker.addEventListener("change", () => {
        const val = picker.value;
        if (val === "__custom__") {
            setLogoUrl("", { showCustomInput: true });
        } else {
            setLogoUrl(val);
        }
        updateBuilderPreview();
    });
}

// Sets the logo URL input/preview and picks the matching dropdown option (or
// "Custom URL…" if the URL doesn't match any preloaded logo).
function setLogoUrl(url, opts) {
    const picker = document.getElementById("f_logoPicker");
    const urlInput = document.getElementById("f_logoUrl");
    const forceCustom = opts && opts.showCustomInput;

    urlInput.value = url || "";

    const matchingOption = Array.from(picker.options).find(o => o.value === url);
    if (!forceCustom && url && matchingOption) {
        picker.value = url;
        urlInput.classList.add("hidden");
    } else if (url || forceCustom) {
        picker.value = "__custom__";
        urlInput.classList.remove("hidden");
    } else {
        picker.value = "";
        urlInput.classList.add("hidden");
    }
    updateLogoPreview();
}

function updateLogoPreview() {
    const preview = document.getElementById("logoPreview");
    const url = document.getElementById("f_logoUrl").value.trim();
    if (url) {
        preview.src = url;
        preview.classList.remove("hidden");
    } else {
        preview.classList.add("hidden");
    }
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
    urlInput.placeholder = entry.platform === "custom" ? "Profile URL" : "Username";
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
        urlInput.placeholder = select.value === "custom" ? "Profile URL" : "Username";
        updateBuilderPreview();
    });
    urlInput.addEventListener("input", updateBuilderPreview);
    customIconInput.addEventListener("input", updateBuilderPreview);
    moveUpBtn.addEventListener("click", () => {
        const prev = row.previousElementSibling;
        if (prev) {
            row.parentNode.insertBefore(row, prev);
            updateBuilderPreview();
        }
    });
    moveDownBtn.addEventListener("click", () => {
        const next = row.nextElementSibling;
        if (next) {
            row.parentNode.insertBefore(next, row);
            updateBuilderPreview();
        }
    });
    removeBtn.addEventListener("click", () => {
        row.remove();
        socialLinks = socialLinks.filter(s => s.id !== id);
        updateBuilderPreview();
    });

    row.appendChild(select);
    row.appendChild(urlInput);
    row.appendChild(customIconInput);
    row.appendChild(moveUpBtn);
    row.appendChild(moveDownBtn);
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

    ["tg_logo", "tg_email", "tg_phone", "tg_website", "tg_social", "tg_tagline"].forEach(id => {
        document.getElementById(id).addEventListener("change", () => {
            updateConditionalVisibility();
            updateBuilderPreview();
        });
    });

    ["f_name", "f_title", "f_company", "f_email", "f_phone", "f_website", "f_taglineText", "f_taglineUrl", "f_logoUrl"]
        .forEach(id => document.getElementById(id).addEventListener("input", () => {
            if (id === "f_logoUrl") updateLogoPreview();
            updateBuilderPreview();
        }));

    document.getElementById("scaleDownBtn").addEventListener("click", () => adjustScale(-10));
    document.getElementById("scaleUpBtn").addEventListener("click", () => adjustScale(10));
    document.getElementById("spacingDownBtn").addEventListener("click", () => adjustSpacing(-10));
    document.getElementById("spacingUpBtn").addEventListener("click", () => adjustSpacing(10));

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

function adjustScale(delta) {
    scalePercent = clamp(scalePercent + delta, 50, 200);
    document.getElementById("scaleValue").textContent = scalePercent + "%";
    updateBuilderPreview();
}

function adjustSpacing(delta) {
    spacingPercent = clamp(spacingPercent + delta, 50, 200);
    document.getElementById("spacingValue").textContent = spacingPercent + "%";
    updateBuilderPreview();
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
            showStatus("At least enter a name for the builder signature.", true);
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
