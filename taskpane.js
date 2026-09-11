/* Liyatex CRM Outlook eklentisi — görev paneli.
   Açık e-postadan görev / toplantı / fuar notu açar, isteğe bağlı e-postayı email aktivitesi olarak kaydeder.
   Kimlik: MSAL (kullanıcının liyatexcrm hesabı) → Dataverse Web API, doğrudan tarayıcıdan (CORS destekli). */
(function () {
  "use strict";
  var C = window.LIY_CFG;
  var API = C.CRM_URL + "/api/data/v9.2/";
  var SCOPES = [C.CRM_URL + "/user_impersonation"];
  var $ = function (id) { return document.getElementById(id); };

  var state = {
    token: null, expiresOn: 0, user: null, userId: null,
    mail: null, account: null, contacts: [], tab: "task",
    lists: { taskTypes: [], fairs: [], seasons: [] }
  };

  var msal = new window.msal.PublicClientApplication({
    auth: { clientId: C.CLIENT_ID, authority: "https://login.microsoftonline.com/" + C.TENANT_ID, redirectUri: C.HOST_URL + "/auth.html" },
    cache: { cacheLocation: "localStorage" }
  });

  /* ---------- kimlik ---------- */
  function haveToken() { return state.token && Date.now() < state.expiresOn - 60000; }

  function silentToken() {
    return msal.initialize().then(function () {
      var acc = msal.getAllAccounts()[0];
      if (!acc) throw new Error("no-account");
      return msal.acquireTokenSilent({ scopes: SCOPES, account: acc });
    }).then(function (r) { setToken(r.accessToken, r.expiresOn, r.account); });
  }

  function setToken(tok, exp, acc) {
    state.token = tok;
    state.expiresOn = exp ? new Date(exp).getTime() : Date.now() + 50 * 60000;
    state.user = acc ? { username: acc.username, name: acc.name } : state.user;
    $("user").textContent = state.user ? (state.user.name || state.user.username) : "";
    $("user").title = state.user ? state.user.username : "";
  }

  function dialogToken() {
    return new Promise(function (resolve, reject) {
      Office.context.ui.displayDialogAsync(C.HOST_URL + "/auth.html", { height: 60, width: 30, promptBeforeOpen: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error("Giriş penceresi açılamadı: " + res.error.message)); return; }
        var dlg = res.value;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var m; try { m = JSON.parse(arg.message); } catch (e) { m = { ok: false, error: "geçersiz yanıt" }; }
          dlg.close();
          if (m.ok) { setToken(m.token, m.expiresOn, { username: m.username, name: m.name }); resolve(); }
          else reject(new Error(m.error || "Giriş başarısız"));
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
          if (arg.error === 12006) reject(new Error("Giriş penceresi kapatıldı."));
        });
      });
    });
  }

  function ensureToken() {
    if (haveToken()) return Promise.resolve();
    return silentToken().catch(function () { return dialogToken(); });
  }

  /* ---------- Dataverse ---------- */
  function api(method, path, body, extraHeaders) {
    return ensureToken().then(function () {
      var h = {
        "Authorization": "Bearer " + state.token, "Accept": "application/json",
        "OData-MaxVersion": "4.0", "OData-Version": "4.0", "Content-Type": "application/json; charset=utf-8"
      };
      if (extraHeaders) Object.keys(extraHeaders).forEach(function (k) { h[k] = extraHeaders[k]; });
      return fetch(API + path, { method: method, headers: h, body: body ? JSON.stringify(body) : undefined });
    }).then(function (r) {
      if (r.status === 204) {
        var loc = r.headers.get("OData-EntityId") || "";
        var m = /\(([0-9a-f-]{36})\)/i.exec(loc);
        return { id: m ? m[1] : null };
      }
      return r.text().then(function (t) {
        var j = {}; try { j = t ? JSON.parse(t) : {}; } catch (e) { j = { raw: t }; }
        if (!r.ok) {
          var msg = (j.error && j.error.message) || (r.status + " " + r.statusText);
          if (r.status === 401) { state.token = null; }
          throw new Error(msg);
        }
        return j;
      });
    });
  }
  function q(v) { return String(v || "").replace(/'/g, "''"); }

  /* ---------- UI yardımcıları ---------- */
  function showMsg(kind, text, link) {
    var el = $("msg"); el.className = "msg " + kind; el.textContent = "";
    el.appendChild(document.createTextNode(text));
    if (link) {
      el.appendChild(document.createTextNode(" "));
      var a = document.createElement("a"); a.href = link.href; a.target = "_blank"; a.rel = "noopener"; a.textContent = link.text; el.appendChild(a);
    }
  }
  function recordUrl(etn, id) {
    return C.CRM_URL + "/main.aspx?appid=" + C.APP_ID + "&pagetype=entityrecord&etn=" + etn + "&id=" + id;
  }
  function fillSelect(sel, items, keepFirst) {
    while (sel.options.length > (keepFirst ? 1 : 0)) sel.remove(sel.options.length - 1);
    items.forEach(function (it) { var o = document.createElement("option"); o.value = it.value; o.textContent = it.label; sel.appendChild(o); });
  }
  function localDT(d) { // datetime-local için yerel biçim
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "T" + p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function isoFromLocal(v) { return v ? new Date(v).toISOString() : null; }
  function isoFromDate(v, hour) { if (!v) return null; var d = new Date(v + "T" + (hour || "09:00") + ":00"); return d.toISOString(); }

  /* ---------- arama kutuları (firma / kişi) ---------- */
  function lookup(inputId, listId, search, onPick) {
    var inp = $(inputId), ul = $(listId), timer = null, items = [];
    function close() { ul.hidden = true; ul.textContent = ""; items = []; }
    function render() {
      ul.textContent = "";
      if (!items.length) { close(); return; }
      items.forEach(function (it, i) {
        var li = document.createElement("li"); li.setAttribute("role", "option"); li.dataset.i = i;
        li.appendChild(document.createTextNode(it.label));
        if (it.sub) { var s = document.createElement("small"); s.textContent = it.sub; li.appendChild(s); }
        li.addEventListener("mousedown", function (e) { e.preventDefault(); onPick(it); inp.value = ""; close(); });
        ul.appendChild(li);
      });
      ul.hidden = false;
    }
    inp.addEventListener("input", function () {
      clearTimeout(timer);
      var v = inp.value.trim();
      if (v.length < 2) { close(); return; }
      timer = setTimeout(function () {
        search(v).then(function (res) { items = res; render(); }).catch(function (e) { showMsg("err", "Arama hatası: " + e.message); });
      }, 250);
    });
    inp.addEventListener("blur", function () { setTimeout(close, 150); });
    inp.addEventListener("keydown", function (e) { if (e.key === "Escape") close(); });
  }

  function searchAccounts(v) {
    return api("GET", "accounts?$select=name,accountid,emailaddress1&$filter=statecode eq 0 and contains(name,'" + q(v) + "')&$orderby=name&$top=8")
      .then(function (r) { return r.value.map(function (a) { return { id: a.accountid, label: a.name, sub: a.emailaddress1 || "" }; }); });
  }
  function searchContacts(v) {
    var f = "statecode eq 0 and contains(fullname,'" + q(v) + "')";
    if (state.account) f += " and (_parentcustomerid_value eq " + state.account.id + " or zeno_brandid/_zeno_accountid_value eq " + state.account.id + ")";
    return api("GET", "contacts?$select=fullname,contactid,emailaddress1,jobtitle&$expand=parentcustomerid_account($select=name)&$filter=" + f + "&$orderby=fullname&$top=8")
      .then(function (r) {
        return r.value.map(function (c) {
          var firm = c.parentcustomerid_account ? c.parentcustomerid_account.name : "";
          return { id: c.contactid, label: c.fullname, sub: [c.jobtitle, firm, c.emailaddress1].filter(Boolean).join(" · "), email: c.emailaddress1 };
        });
      });
  }

  function setAccount(a) {
    state.account = a; var box = $("accChip"); box.textContent = "";
    if (!a) return;
    var ch = document.createElement("span"); ch.className = "chip"; ch.appendChild(document.createTextNode(a.label));
    var x = document.createElement("button"); x.type = "button"; x.setAttribute("aria-label", "Firmayı kaldır"); x.textContent = "×";
    x.addEventListener("click", function () { setAccount(null); }); ch.appendChild(x); box.appendChild(ch);
  }
  function addContact(c) {
    if (state.contacts.some(function (x) { return x.id === c.id; })) return;
    state.contacts.push(c); renderContacts();
  }
  function renderContacts() {
    var box = $("conChips"); box.textContent = "";
    state.contacts.forEach(function (c) {
      var ch = document.createElement("span"); ch.className = "chip"; ch.appendChild(document.createTextNode(c.label));
      var x = document.createElement("button"); x.type = "button"; x.setAttribute("aria-label", "Kişiyi kaldır"); x.textContent = "×";
      x.addEventListener("click", function () { state.contacts = state.contacts.filter(function (y) { return y.id !== c.id; }); renderContacts(); });
      ch.appendChild(x); box.appendChild(ch);
    });
  }

  /* ---------- e-postadan otomatik eşleme ---------- */
  function autoMatch() {
    var m = state.mail; if (!m || !m.fromEmail) return Promise.resolve();
    var email = m.fromEmail.toLowerCase(), domain = email.split("@")[1] || "";
    var hint = [];
    return api("GET", "contacts?$select=fullname,contactid,emailaddress1,jobtitle,_parentcustomerid_value&$expand=parentcustomerid_account($select=name,accountid)&$filter=statecode eq 0 and (emailaddress1 eq '" + q(email) + "' or emailaddress2 eq '" + q(email) + "')&$top=3")
      .then(function (r) {
        if (r.value.length) {
          var c = r.value[0];
          addContact({ id: c.contactid, label: c.fullname, email: c.emailaddress1 });
          hint.push("Gönderen CRM'de kayıtlı kişi: " + c.fullname + ".");
          if (c.parentcustomerid_account) { setAccount({ id: c.parentcustomerid_account.accountid, label: c.parentcustomerid_account.name }); }
          if (r.value.length > 1) hint.push("Aynı adresle " + r.value.length + " kişi var, ilki seçildi.");
        }
        if (state.account || !domain || /^(gmail|hotmail|outlook|yahoo|icloud|yandex)\./.test(domain)) return null;
        return api("GET", "accounts?$select=name,accountid&$filter=statecode eq 0 and (contains(emailaddress1,'@" + q(domain) + "') or contains(websiteurl,'" + q(domain) + "'))&$top=3");
      })
      .then(function (r) {
        if (r && r.value.length === 1) { setAccount({ id: r.value[0].accountid, label: r.value[0].name }); hint.push("Firma alan adından bulundu: " + r.value[0].name + "."); }
        else if (r && r.value.length > 1) hint.push("Alan adı " + r.value.length + " firmayla eşleşti, elle seçin.");
        if (!state.account && !state.contacts.length) hint.push("Gönderen CRM'de bulunamadı; firmayı elle seçin.");
        $("matchHint").textContent = hint.join(" ");
      })
      .then(checkExistingEmail);
  }

  function checkExistingEmail() {
    var mid = state.mail.messageId; if (!mid) return;
    return api("GET", "emails?$select=activityid,subject&$filter=messageid eq '" + q(mid) + "'&$top=1").then(function (r) {
      if (r.value.length) {
        state.mail.existingEmailId = r.value[0].activityid;
        $("mLinked").textContent = "Bu e-posta CRM'de zaten kayıtlı; tekrar kaydedilmez.";
        $("saveMail").checked = false;
      }
    }).catch(function () { /* bilgi amaçlı, sessiz geç */ });
  }

  /* ---------- referans listeleri ---------- */
  function loadLists() {
    return Promise.all([
      api("GET", "EntityDefinitions(LogicalName='task')/Attributes(LogicalName='zeno_gorevturu')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"),
      api("GET", "zeno_exhibitions?$select=zeno_name,zeno_exhibitionid&$filter=statecode eq 0&$orderby=createdon desc&$top=12"),
      api("GET", "EntityDefinitions(LogicalName='appointment')/Attributes(LogicalName='zeno_sezon')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"),
      api("GET", "WhoAmI")
    ]).then(function (res) {
      var lbl = function (o) { return (o.Label.UserLocalizedLabel || {}).Label || String(o.Value); };
      fillSelect($("tType"), res[0].OptionSet.Options.map(function (o) { return { value: o.Value, label: lbl(o) }; }), true);
      fillSelect($("eFair"), res[1].value.map(function (f) { return { value: f.zeno_exhibitionid, label: f.zeno_name }; }), true);
      fillSelect($("aSeason"), res[2].OptionSet.Options.map(function (o) { return { value: o.Value, label: lbl(o) }; }), true);
      state.userId = res[3].UserId;
      if (res[1].value.length) $("eFair").value = res[1].value[0].zeno_exhibitionid;
    });
  }

  /* ---------- e-postayı oku ---------- */
  function readMail() {
    var it = Office.context.mailbox.item;
    var m = {
      subject: it.subject || "", fromName: it.from ? it.from.displayName : "", fromEmail: it.from ? it.from.emailAddress : "",
      to: (it.to || []).map(function (r) { return { name: r.displayName, email: r.emailAddress }; }),
      cc: (it.cc || []).map(function (r) { return { name: r.displayName, email: r.emailAddress }; }),
      date: it.dateTimeCreated ? new Date(it.dateTimeCreated) : new Date(),
      messageId: it.internetMessageId || "", itemId: it.itemId || "", text: "", html: ""
    };
    state.mail = m;
    $("mSubject").textContent = m.subject || "(konu yok)";
    $("mFrom").textContent = m.fromName ? m.fromName + " <" + m.fromEmail + ">" : m.fromEmail;
    $("mDate").textContent = m.date.toLocaleString("tr-TR");
    $("subject").value = m.subject;
    return new Promise(function (resolve) {
      it.body.getAsync(Office.CoercionType.Text, function (r) {
        m.text = r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : "";
        $("desc").value = buildDescription(m);
        it.body.getAsync(Office.CoercionType.Html, function (r2) {
          m.html = r2.status === Office.AsyncResultStatus.Succeeded ? (r2.value || "") : "";
          resolve();
        });
      });
    });
  }
  function buildDescription(m) {
    var t = (m.text || "").replace(/\r/g, "").replace(/\n{3,}/g, "\n\n").trim();
    if (t.length > C.BODY_LIMIT) t = t.slice(0, C.BODY_LIMIT) + "\n[… kısaltıldı]";
    return "Kaynak e-posta: " + (m.fromName || m.fromEmail) + " — " + m.date.toLocaleString("tr-TR") + "\n\n" + t;
  }

  /* ---------- kayıt oluşturma ---------- */
  function regardingBind(entityLogical) {
    // Aktivite tablolarında regardingobjectid gezinme adı: regardingobjectid_<hedef>_<tablo>
    if (state.account) return { key: "regardingobjectid_account_" + entityLogical + "@odata.bind", val: "/accounts(" + state.account.id + ")" };
    if (state.contacts.length) return { key: "regardingobjectid_contact_" + entityLogical + "@odata.bind", val: "/contacts(" + state.contacts[0].id + ")" };
    return null;
  }
  function contactParties(mask) {
    return state.contacts.map(function (c) { return { "partyid_contact@odata.bind": "/contacts(" + c.id + ")", participationtypemask: mask }; });
  }

  function buildTask() {
    var b = { subject: $("subject").value.trim(), description: $("desc").value, prioritycode: Number($("tPrio").value), zeno_kaynakturu: 100000003 };
    if ($("tType").value) b.zeno_gorevturu = Number($("tType").value);
    var due = isoFromDate($("tDue").value, "18:00"); if (due) b.scheduledend = due;
    var r = regardingBind("task"); if (r) b[r.key] = r.val;
    return { etn: "task", set: "tasks", body: b, label: "Görev" };
  }
  function buildAppt() {
    var s = isoFromLocal($("aStart").value), e = isoFromLocal($("aEnd").value);
    if (!s || !e) throw new Error("Toplantı başlangıç ve bitiş saati gerekli.");
    if (new Date(e) <= new Date(s)) throw new Error("Bitiş, başlangıçtan sonra olmalı.");
    var b = { subject: $("subject").value.trim(), description: $("desc").value, scheduledstart: s, scheduledend: e, location: $("aLoc").value.trim() || null, zeno_kaynakturu: 100000003 };
    if ($("aSeason").value) b.zeno_sezon = Number($("aSeason").value);
    if (state.userId) b["zeno_TakenById_Appointment@odata.bind"] = "/systemusers(" + state.userId + ")";
    var r = regardingBind("appointment"); if (r) b[r.key] = r.val;
    if (state.contacts.length) {
      b.appointment_activity_parties = contactParties(5); // 5 = zorunlu katılımcı
      b.zeno_customerattendees = state.contacts.map(function (c) { return c.label; }).join("\n");
    }
    return { etn: "appointment", set: "appointments", body: b, label: "Toplantı" };
  }
  function buildExpo() {
    if (!$("eFair").value) throw new Error("Fuar seçin.");
    var b = { subject: $("subject").value.trim(), description: $("desc").value, zeno_interactiontypecode: Number($("eKind").value) };
    b["zeno_exhibitionid_zeno_exhibitionnote@odata.bind"] = "/zeno_exhibitions(" + $("eFair").value + ")";
    var d = isoFromDate($("eDate").value, "10:00"); if (d) b.scheduledstart = d;
    if (state.userId) b["zeno_visitorid_zeno_exhibitionnote@odata.bind"] = "/systemusers(" + state.userId + ")";
    var r = regardingBind("zeno_exhibitionnote"); if (r) b[r.key] = r.val;
    if (state.contacts.length) {
      b.zeno_exhibitionnote_activity_parties = contactParties(5);
      b["zeno_ContactId_zeno_exhibitionnote@odata.bind"] = "/contacts(" + state.contacts[0].id + ")";
    }
    return { etn: "zeno_exhibitionnote", set: "zeno_exhibitionnotes", body: b, label: "Fuar notu" };
  }

  function saveEmailActivity() {
    var m = state.mail;
    if (!$("saveMail").checked || m.existingEmailId) return Promise.resolve(null);
    var parties = [];
    var sender = state.contacts.filter(function (c) { return (c.email || "").toLowerCase() === m.fromEmail.toLowerCase(); })[0];
    if (sender) parties.push({ "partyid_contact@odata.bind": "/contacts(" + sender.id + ")", participationtypemask: 1 });
    else parties.push({ addressused: m.fromEmail, participationtypemask: 1 });
    m.to.forEach(function (r) { parties.push({ addressused: r.email, participationtypemask: 2 }); });
    m.cc.forEach(function (r) { parties.push({ addressused: r.email, participationtypemask: 3 }); });
    var html = m.html || ""; if (html.length > 200000) html = html.slice(0, 200000) + "<p>[… kısaltıldı]</p>";
    var b = {
      subject: m.subject || "(konu yok)", description: html, directioncode: false, messageid: m.messageId || null,
      sender: m.fromEmail, torecipients: m.to.map(function (r) { return r.email; }).join(";"),
      actualstart: m.date.toISOString(), actualend: m.date.toISOString(), email_activity_parties: parties
    };
    var r = regardingBind("email"); if (r) b[r.key] = r.val;
    return api("POST", "emails", b).then(function (res) {
      // Gelen posta olarak kapat (statecode 1 / statuscode 4 = Alındı). Onay eklentisi yalnız zeno_approvalstatus=1 olanı durdurur.
      return api("PATCH", "emails(" + res.id + ")", { statecode: 1, statuscode: 4 }).then(function () { return res.id; }, function () { return res.id; });
    });
  }

  function onSave() {
    var btn = $("btnSave"); btn.disabled = true; showMsg("", "");
    var spec;
    try {
      if (!$("subject").value.trim()) throw new Error("Konu boş olamaz.");
      spec = state.tab === "task" ? buildTask() : state.tab === "appt" ? buildAppt() : buildExpo();
    } catch (e) { showMsg("err", e.message); btn.disabled = false; return; }
    var created = null;
    api("POST", spec.set, spec.body)
      .then(function (r) { created = r.id; return saveEmailActivity(); })
      .then(function (emailId) {
        var extra = emailId ? " E-posta da kaydedildi." : "";
        showMsg("ok", spec.label + " oluşturuldu." + extra, { href: recordUrl(spec.etn, created), text: "CRM'de aç" });
      })
      .catch(function (e) {
        if (created) showMsg("err", spec.label + " oluşturuldu ama e-posta kaydedilemedi: " + e.message, { href: recordUrl(spec.etn, created), text: "CRM'de aç" });
        else showMsg("err", "Kayıt oluşturulamadı: " + e.message);
      })
      .then(function () { btn.disabled = false; });
  }

  /* ---------- sekmeler ---------- */
  function selectTab(name) {
    state.tab = name;
    document.querySelectorAll(".tabs [role=tab]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.tab === name)); });
    $("pTask").hidden = name !== "task"; $("pAppt").hidden = name !== "appt"; $("pExpo").hidden = name !== "expo";
  }

  /* ---------- başlat ---------- */
  function boot() {
    document.querySelectorAll(".tabs [role=tab]").forEach(function (b) { b.addEventListener("click", function () { selectTab(b.dataset.tab); }); });
    lookup("accQ", "accList", searchAccounts, setAccount);
    lookup("conQ", "conList", searchContacts, addContact);
    $("btnSave").addEventListener("click", onSave);
    $("btnLogin").addEventListener("click", function () {
      $("loginMsg").className = "msg";
      dialogToken().then(start).catch(function (e) { $("loginMsg").className = "msg err"; $("loginMsg").textContent = e.message; });
    });
    // varsayılan tarihler
    var d = new Date(); d.setDate(d.getDate() + 3); $("tDue").value = d.toISOString().slice(0, 10);
    var s = new Date(); s.setDate(s.getDate() + 1); s.setHours(10, 0, 0, 0); $("aStart").value = localDT(s);
    var e = new Date(s.getTime() + 60 * 60000); $("aEnd").value = localDT(e);
    $("eDate").value = new Date().toISOString().slice(0, 10);

    silentToken().then(start, function () { $("login").style.display = "block"; });
  }

  function start() {
    $("login").style.display = "none"; $("app").hidden = false;
    readMail().then(loadLists).then(autoMatch).catch(function (e) { showMsg("err", e.message); });
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) { document.body.textContent = "Bu eklenti yalnız Outlook'ta çalışır."; return; }
    boot();
  });
})();
