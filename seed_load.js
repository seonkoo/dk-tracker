window.SEED_DATA = JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(window.__SEED_B64), c => c.charCodeAt(0))));
