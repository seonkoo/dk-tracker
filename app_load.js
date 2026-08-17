document.open();document.write(new TextDecoder().decode(Uint8Array.from(atob(window.__APP_B64), c => c.charCodeAt(0))));document.close();
