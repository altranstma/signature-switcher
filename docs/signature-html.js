/* global Office */
/*
 * Shared by taskpane.js (preview/edit) and launchevent.js (signature switching).
 * Builder-mode entries store only `fields`, not rendered `html` — see onSave()
 * in taskpane.js for why — so both call sites need to render HTML the same way.
 */

const SOCIAL_ICON_BASE = "https://altranstma.github.io/signature-switcher/assets/social";
const SOCIAL_PLATFORMS = {
    instagram: { label: "Instagram", icon: `${SOCIAL_ICON_BASE}/instagram.png` },
    facebook: { label: "Facebook", icon: `${SOCIAL_ICON_BASE}/facebook.png` },
    x: { label: "X (Twitter)", icon: `${SOCIAL_ICON_BASE}/x.png` },
    linkedin: { label: "LinkedIn", icon: `${SOCIAL_ICON_BASE}/linkedin.png` },
    youtube: { label: "YouTube", icon: `${SOCIAL_ICON_BASE}/youtube.png` },
    custom: { label: "Custom", icon: null }
};

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

    const socialImgs = (fields.socialLinks || [])
        .filter(s => s.url)
        .map(s => {
            const icon = s.platform === "custom" ? sanitizeHttpUrl(s.customIconUrl) : (SOCIAL_PLATFORMS[s.platform] || {}).icon;
            const linkUrl = sanitizeHttpUrl(s.url);
            if (!icon || !linkUrl) return "";
            return `<a href="${linkUrl}" style="text-decoration:none;margin-right:6px;"><img src="${icon}" width="20" height="20" style="border:0;vertical-align:middle;" alt="" /></a>`;
        })
        .join("");

    let html = `<table cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif;border-collapse:collapse;"><tr>`;
    html += logoCell;
    html += `<td style="${fields.logoUrl ? "border-left:2px solid #cccccc;padding-left:14px;" : ""}vertical-align:middle;">`;
    if (fields.name) html += `<div style="font-size:14px;font-weight:bold;color:#222222;">${escapeHtml(fields.name)}</div>`;
    if (titleCompany) html += `<div style="font-size:12px;font-weight:bold;color:#666666;">${escapeHtml(titleCompany)}</div>`;
    if (contactLine) html += `<div style="font-size:12px;color:#333333;margin-top:3px;">${contactLine}</div>`;
    if (socialImgs) html += `<div style="margin-top:6px;">${socialImgs}</div>`;
    html += `</td></tr>`;
    if (fields.tagline && fields.tagline.text) {
        const taglineUrl = sanitizeHttpUrl(fields.tagline.url);
        const taglineContent = taglineUrl
            ? `<a href="${taglineUrl}" style="color:#1155cc;text-decoration:underline;">${escapeHtml(fields.tagline.text)}</a>`
            : escapeHtml(fields.tagline.text);
        html += `<tr><td colspan="2" style="padding-top:8px;font-size:11px;color:#555555;">${taglineContent}</td></tr>`;
    }
    html += `</table>`;
    return html;
}

// Resolves any stored map entry (builder, custom, or legacy plain-string) to
// renderable signature HTML. Builder entries render fields on demand instead
// of storing a redundant html copy, since Office.context.roamingSettings has
// a hard 32KB-per-user cap and duplicating html+fields burned through it fast.
function resolveSignatureHtml(entry) {
    if (!entry) return "";
    if (typeof entry === "string") return entry;
    if (entry.mode === "custom") return entry.html || "";
    if (entry.fields) return buildSignatureHtml(entry.fields);
    return entry.html || "";
}
