/* Service worker for Maison Retrouvailles – viser push-varsel når drinken er klar */
try { importScripts("config.js"); } catch (e) { }
var API = (self.BAR_CONFIG && self.BAR_CONFIG.api) || "";
var TOKEN = new URL(self.location.href).searchParams.get("g") || "";

self.addEventListener("install", function () { self.skipWaiting(); });
self.addEventListener("activate", function (e) { e.waitUntil(self.clients.claim()); });

function hentTekst() {
  var standard = { tittel: "Votre verre est prêt ✨", tekst: "Drinken din venter på deg ved baren." };
  if (!API || !TOKEN) return Promise.resolve(standard);
  var ctrl = new AbortController();
  var t = setTimeout(function () { ctrl.abort(); }, 4500);
  var url = API + "?a=varsel&d=" + encodeURIComponent(JSON.stringify({ token: TOKEN })) + "&_=" + Date.now();
  return fetch(url, { credentials: "omit", cache: "no-store", signal: ctrl.signal })
    .then(function (r) { return r.json(); })
    .then(function (j) { clearTimeout(t); return (j && j.tittel) ? j : standard; })
    .catch(function () { clearTimeout(t); return standard; });
}

self.addEventListener("push", function (e) {
  e.waitUntil(hentTekst().then(function (v) {
    return self.registration.showNotification(v.tittel, {
      body: v.tekst,
      icon: "ikoner/ikon-192.png",
      badge: "ikoner/ikon-192.png",
      tag: "retrouvailles-" + Date.now(),
      data: { url: "./?g=" + encodeURIComponent(TOKEN) }
    });
  }).then(function () {
    return self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (liste) {
      liste.forEach(function (c) { c.postMessage({ type: "push" }); });
    });
  }));
});

self.addEventListener("notificationclick", function (e) {
  e.notification.close();
  var url = (e.notification.data && e.notification.data.url) || "./";
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (liste) {
    for (var i = 0; i < liste.length; i++) if ("focus" in liste[i]) return liste[i].focus();
    return self.clients.openWindow(url);
  }));
});
