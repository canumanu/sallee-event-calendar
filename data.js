
/* Sallee Event Calendar — data layer.
   SharePointStore: Microsoft 365 sign-in (MSAL) + Microsoft Graph, data in a SharePoint list.
   DemoStore: sample events from events.json, kept in memory (used when config.js has no clientId). */
(function () {
  const C = window.SALLEE_CAL_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["User.Read", "Sites.ReadWrite.All"];
  const SETUP_SCOPES = ["Sites.Manage.All"];

  // App values <-> SharePoint choice labels
  const TYPE_TO_SP = { raceday: "Race day / stakes", sale: "Sale", meet: "Race meet", tack: "Under tack show", digital: "Digital sale", recruiting: "Recruiting" };
  const STATUS_TO_SP = { planned: "Planned", scheduled: "Shoot booked", shot: "Filmed", posted: "Posted" };
  const invert = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]));
  const SP_TO_TYPE = invert(TYPE_TO_SP), SP_TO_STATUS = invert(STATUS_TO_SP);

  // App field -> SharePoint internal column name
  const FIELD = {
    title: "Title", venue: "EventVenue", category: "EventType", start: "EventStart", end: "EventEnd",
    location: "EventLocation", notes: "EventNotes", marquee: "Marquee", status: "FootageStatus",
    tbc: "DatesTBC", shots: "ShotsJson", copiedFrom: "CopiedFrom"
  };
  const SELECT = Object.values(FIELD).join(",");

  function toSP(patch) {
    const f = {};
    for (const [k, v] of Object.entries(patch)) {
      const n = FIELD[k]; if (!n) continue;
      if (k === "category") f[n] = TYPE_TO_SP[v] || v;
      else if (k === "status") f[n] = STATUS_TO_SP[v] || v;
      else if (k === "shots") f[n] = JSON.stringify(v || []);
      else if (k === "marquee" || k === "tbc") f[n] = !!v;
      else f[n] = v == null ? "" : String(v);
    }
    return f;
  }
  function fromSP(item) {
    const f = item.fields || {};
    let shots = [];
    try { shots = JSON.parse(f.ShotsJson || "[]"); } catch (e) { shots = []; }
    const start = (f.EventStart || "").trim();
    return {
      id: String(item.id), title: f.Title || "", venue: f.EventVenue || "",
      category: SP_TO_TYPE[f.EventType] || "sale", start, end: (f.EventEnd || "").trim() || start,
      location: f.EventLocation || "", notes: f.EventNotes || "", marquee: !!f.Marquee,
      status: SP_TO_STATUS[f.FootageStatus] || "planned", tbc: !!f.DatesTBC,
      shots: Array.isArray(shots) ? shots : [], copiedFrom: f.CopiedFrom || undefined
    };
  }

  /* ------------------------------------------------------------------ */
  function SharePointStore() {
    let pca, account, siteId, listId, listUrl = "", recipientsUrl = "", canWrite = null;

    async function initAuth() {
      pca = new msal.PublicClientApplication({
        auth: {
          clientId: C.clientId,
          authority: "https://login.microsoftonline.com/" + (C.tenantId || "organizations"),
          redirectUri: location.origin + location.pathname
        },
        cache: { cacheLocation: "localStorage" }
      });
      await pca.initialize();
      const res = await pca.handleRedirectPromise();
      account = (res && res.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
      if (account) pca.setActiveAccount(account);
      return account;
    }
    async function token(scopes) {
      try { return (await pca.acquireTokenSilent({ scopes, account })).accessToken; }
      catch (e) {
        if (e instanceof msal.InteractionRequiredAuthError) {
          await pca.acquireTokenRedirect({ scopes, account });
          return new Promise(() => {}); // page is navigating away
        }
        throw e;
      }
    }
    async function graph(method, path, body, scopes = SCOPES) {
      const t = await token(scopes);
      const r = await fetch(path.startsWith("http") ? path : GRAPH + path, {
        method,
        headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : undefined
      });
      if (!r.ok) {
        const err = new Error("Microsoft Graph returned " + r.status);
        err.status = r.status;
        try { err.detail = await r.json(); } catch (e) {}
        throw err;
      }
      return r.status === 204 ? null : r.json();
    }
    async function findList(name) {
      let url = `/sites/${siteId}/lists?$select=id,displayName,webUrl&$top=200`;
      while (url) {
        const res = await graph("GET", url);
        const hit = (res.value || []).find(l => (l.displayName || "").toLowerCase() === name.toLowerCase());
        if (hit) return hit;
        url = res["@odata.nextLink"] || null;
      }
      return null;
    }
    // Optional: ask SharePoint whether this person may edit the list. Needs the SharePoint
    // "AllSites.Read" delegated permission (README step 1). Without it we learn from the first save.
    async function checkEditRights() {
      try {
        const r = await pca.acquireTokenSilent({ scopes: [`https://${C.siteHost}/AllSites.Read`], account });
        const url = `https://${C.siteHost}${C.sitePath}/_api/web/lists/getbytitle('${C.eventsList.replace(/'/g, "''")}')/EffectiveBasePermissions`;
        const resp = await fetch(url, { headers: { Authorization: "Bearer " + r.accessToken, Accept: "application/json;odata=nometadata" } });
        if (!resp.ok) return null;
        const p = await resp.json();
        return (Number(p.Low) & 0x4) !== 0; // EditListItems
      } catch (e) { return null; }
    }

    return {
      mode: "sharepoint",
      get canWrite() { return canWrite; },
      set canWrite(v) { canWrite = v; },
      get listUrl() { return listUrl; },
      get recipientsUrl() { return recipientsUrl; },
      get user() { return account ? (account.name || account.username) : ""; },

      async init() {
        await initAuth();
        if (!account) return "signedOut";
        const site = await graph("GET", `/sites/${C.siteHost}:${C.sitePath}?$select=id,webUrl`);
        siteId = site.id;
        const list = await findList(C.eventsList);
        const rec = await findList(C.recipientsList).catch(() => null);
        if (rec) recipientsUrl = rec.webUrl;
        if (!list) return "setup";
        listId = list.id; listUrl = list.webUrl;
        canWrite = await checkEditRights();
        return "ready";
      },
      signIn() { return pca.loginRedirect({ scopes: SCOPES }); },
      signOut() { return pca.logoutRedirect({ account }); },

      async load() {
        const m = new Map();
        let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=${SELECT})&$top=500`;
        while (url) {
          const res = await graph("GET", url);
          for (const it of res.value || []) { const e = fromSP(it); if (e.start) m.set(e.id, e); }
          url = res["@odata.nextLink"] || null;
        }
        return m;
      },
      async update(id, patch) { await graph("PATCH", `/sites/${siteId}/lists/${listId}/items/${id}/fields`, toSP(patch)); },
      async create(doc) {
        const res = await graph("POST", `/sites/${siteId}/lists/${listId}/items`, { fields: toSP(doc) });
        return String(res.id);
      },
      async remove(id) { await graph("DELETE", `/sites/${siteId}/lists/${listId}/items/${id}`); },

      // One-time: create both lists and import events.json. Needs Sites.Manage.All (site owner).
      async setup(progress) {
        const text = (name, displayName, extra = {}) => ({ name, displayName, text: extra });
        const bool = (name, displayName) => ({ name, displayName, boolean: {} });
        const choice = (name, displayName, choices) => ({ name, displayName, choice: { choices, displayAs: "dropDownMenu" } });
        let list = await findList(C.eventsList);
        if (!list) {
          progress("Creating the “" + C.eventsList + "” list…");
          list = await graph("POST", `/sites/${siteId}/lists`, {
            displayName: C.eventsList,
            list: { template: "genericList" },
            columns: [
              text("EventVenue", "Venue"),
              choice("EventType", "Type", Object.values(TYPE_TO_SP)),
              text("EventStart", "Start (YYYY-MM-DD)", { maxLength: 10 }),
              text("EventEnd", "End (YYYY-MM-DD)", { maxLength: 10 }),
              text("EventLocation", "Location"),
              bool("Marquee", "Marquee"),
              choice("FootageStatus", "Footage status", Object.values(STATUS_TO_SP)),
              bool("DatesTBC", "Dates TBC"),
              text("EventNotes", "Notes", { allowMultipleLines: true, textType: "plain", linesForEditing: 6 }),
              text("ShotsJson", "Shot list (app data)", { allowMultipleLines: true, textType: "plain" }),
              text("CopiedFrom", "Copied from (app data)")
            ]
          }, SETUP_SCOPES);
        }
        listId = list.id; listUrl = list.webUrl;
        let rec = await findList(C.recipientsList);
        if (!rec) {
          progress("Creating the “" + C.recipientsList + "” list…");
          rec = await graph("POST", `/sites/${siteId}/lists`, {
            displayName: C.recipientsList,
            list: { template: "genericList" },
            columns: [
              text("Email", "Email"),
              text("Mobile", "Mobile (+1XXXXXXXXXX)"),
              bool("ByEmail", "Email reminders"),
              bool("ByTeams", "Teams reminders"),
              bool("ByText", "Text reminders"),
              bool("Active", "Active")
            ]
          }, SETUP_SCOPES);
        }
        recipientsUrl = rec.webUrl;
        const existing = await this.load();
        if (existing.size) { progress(""); return existing.size; }
        const seed = await (await fetch("events.json", { cache: "no-store" })).json();
        let n = 0;
        for (const e of seed) {
          await this.create(e);
          n++; progress(`Importing events… ${n} of ${seed.length}`);
        }
        progress("");
        return n;
      }
    };
  }

  /* ------------------------------------------------------------------ */
  function DemoStore() {
    let data = new Map(), n = 0;
    return {
      mode: "demo", canWrite: true, listUrl: "", recipientsUrl: "", user: "",
      async init() {
        const seed = await (await fetch("events.json", { cache: "no-store" })).json();
        for (const e of seed) { const id = "d" + (++n); data.set(id, { ...e, id }); }
        return "ready";
      },
      async load() { return new Map([...data].map(([k, v]) => [k, { ...v }])); },
      async update(id, patch) { data.set(id, { ...data.get(id), ...patch }); },
      async create(doc) { const id = "d" + (++n); data.set(id, { ...doc, id }); return id; },
      async remove(id) { data.delete(id); },
      signIn() {}, signOut() {}
    };
  }

  window.CalendarStore = C.clientId ? SharePointStore() : DemoStore();
})();
