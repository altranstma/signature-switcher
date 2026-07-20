/*
 * Event-based signature switcher.
 * Fires on OnNewMessageCompose (once, at compose start) and OnMessageFromChanged
 * (every time the From address changes). Applies the current user's own
 * per-address signature (stored in Office.context.roamingSettings, managed via
 * taskpane.js) when the active From address is one they've configured.
 * Addresses with no configured signature are left untouched, so Outlook's own
 * native default signature (if any) is never overridden.
 */

const SIGNATURE_MAP_KEY = "signatureMap";
const INJECTED_MARKER_KEY = "sigSwitcherInjected";

function getSignatureMap() {
    const map = Office.context.roamingSettings.get(SIGNATURE_MAP_KEY);
    return map && typeof map === "object" ? map : {};
}

const LOG_PREFIX = "[SigSwitcher]";

function applyForCurrentAddress(event) {
    console.log(LOG_PREFIX, "handler invoked");
    const item = Office.context.mailbox.item;

    item.from.getAsync((fromResult) => {
        console.log(LOG_PREFIX, "item.from.getAsync ->", fromResult.status, fromResult.value);
        if (fromResult.status === Office.AsyncResultStatus.Failed) {
            console.log(LOG_PREFIX, "Could not read From address: " + fromResult.error.message);
            event.completed();
            return;
        }

        const address = (fromResult.value.emailAddress || "").toLowerCase().trim();
        const signatureMap = getSignatureMap();
        console.log(LOG_PREFIX, "address =", address, "signatureMap keys =", Object.keys(signatureMap));
        const entry = signatureMap[address];
        // Entries are {mode, html, fields?} objects; tolerate a legacy plain-string entry too.
        const configuredSignature = entry && (typeof entry === "string" ? entry : entry.html);
        console.log(LOG_PREFIX, "configuredSignature found?", !!configuredSignature, "length =", configuredSignature ? configuredSignature.length : 0);

        item.loadCustomPropertiesAsync((propsResult) => {
            console.log(LOG_PREFIX, "loadCustomPropertiesAsync ->", propsResult.status);
            if (propsResult.status === Office.AsyncResultStatus.Failed) {
                console.log(LOG_PREFIX, "Could not load custom properties: " + propsResult.error.message);
                event.completed();
                return;
            }

            const customProps = propsResult.value;
            const previouslyInjected = customProps.get(INJECTED_MARKER_KEY) === "1";
            console.log(LOG_PREFIX, "previouslyInjected =", previouslyInjected);

            if (configuredSignature) {
                console.log(LOG_PREFIX, "calling setSignatureAsync with configured signature");
                // This address has a user-configured signature: apply it.
                setBodySignature(configuredSignature, () => {
                    customProps.set(INJECTED_MARKER_KEY, "1");
                    customProps.saveAsync((saveResult) => {
                        console.log(LOG_PREFIX, "customProps.saveAsync ->", saveResult.status);
                        event.completed();
                    });
                });
            } else if (previouslyInjected) {
                console.log(LOG_PREFIX, "clearing previously-injected signature");
                // Switched away from a configured address to one with no
                // configuration of its own: clear what we injected rather than
                // leaving a mismatched signature behind.
                setBodySignature("", () => {
                    customProps.set(INJECTED_MARKER_KEY, "0");
                    customProps.saveAsync(() => event.completed());
                });
            } else {
                console.log(LOG_PREFIX, "no configuration and nothing injected - leaving default alone");
                // No configuration for this address and nothing of ours was
                // injected: leave Outlook's own default signature alone.
                event.completed();
            }
        });
    });
}

function setBodySignature(signatureHtml, callback) {
    Office.context.mailbox.item.body.setSignatureAsync(
        signatureHtml,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            console.log(LOG_PREFIX, "setSignatureAsync ->", result.status, result.error ? result.error.message : "");
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(LOG_PREFIX, "setSignatureAsync failed: " + result.error.message);
            }
            callback();
        }
    );
}

function onNewMessageComposeHandler(event) {
    applyForCurrentAddress(event);
}

function onMessageFromChangedHandler(event) {
    applyForCurrentAddress(event);
}

// Maps the manifest's runtime action ids to these handlers. Required in both
// the classic-Windows JS-only runtime and the browser runtime.
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
