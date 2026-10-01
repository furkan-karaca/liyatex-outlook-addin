/* Liyatex CRM Outlook eklentisi — görev paneli.
   Açık e-postadan görev / toplantı / fuar notu açar, isteğe bağlı e-postayı email aktivitesi olarak kaydeder.
   Kimlik: MSAL (kullanıcının liyatexcrm hesabı) → Dataverse Web API, doğrudan tarayıcıdan (CORS destekli). */
(function () {
  "use strict";
  var C = window.LIY_CFG;
  /* Tanı satırı: eski Outlook (Mac) gibi istemcilerde ne olduğunu görmek için — hata ve olay günlüğü panelde. */
  function diag(msg) {
    try { var el = document.getElementById("diag"); if (el) el.textContent += new Date().toLocaleTimeString("tr-TR") + " " + msg + "\n"; } catch (e) { /* yok */ }
  }
  window.addEventListener("error", function (ev) { diag("HATA: " + (ev.message || ev.error)); });
  window.addEventListener("unhandledrejection", function (ev) { diag("HATA(p): " + (ev.reason && (ev.reason.message || ev.reason.errorMessage) || ev.reason)); });
  var API = C.CRM_URL + "/api/data/v9.2/";
  var SCOPES = [C.CRM_URL + "/user_impersonation"];
  var $ = function (id) { return document.getElementById(id); };

  var state = {
    token: null, expiresOn: 0, user: null, userId: null,
    mail: null, account: null, contacts: [], tab: "task",
    lists: { taskTypes: [], fairs: [], seasons: [] },
    /* kayıt düğmesinin durumu — bkz. "MÜKERRER KAYIT" (onSave üstü) */
    busy: false, asking: false, done: null, reloadPending: false, loading: false
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

  var TOK_KEY = "liy_tok";
  function setToken(tok, exp, acc) {
    state.token = tok;
    state.expiresOn = exp ? new Date(exp).getTime() : Date.now() + 50 * 60000;
    state.user = acc ? { username: acc.username, name: acc.name } : state.user;
    $("user").textContent = state.user ? (state.user.name || state.user.username) : "";
    $("user").title = state.user ? state.user.username : "";
    // Outlook masaüstünde iletişim kutusunun depolaması panelden ayrı; jetonu panel tarafında saklayıp süresi dolana dek yeniden sormuyoruz.
    try { localStorage.setItem(TOK_KEY, JSON.stringify({ token: tok, expiresOn: state.expiresOn, user: state.user })); } catch (e) { /* depolama kapalı */ }
  }
  function cachedToken() {
    try {
      var c = JSON.parse(localStorage.getItem(TOK_KEY) || "null");
      if (c && c.token && Date.now() < c.expiresOn - 60000) { state.token = c.token; state.expiresOn = c.expiresOn; state.user = c.user; setToken(c.token, c.expiresOn, c.user); return true; }
    } catch (e) { /* yok say */ }
    return false;
  }

  function dialogToken() {
    return new Promise(function (resolve, reject) {
      var hint = ""; try { var c = JSON.parse(localStorage.getItem(TOK_KEY) || "null"); hint = c && c.user ? c.user.username : ""; } catch (e) { /* yok */ }
      diag("dialog açılıyor");
      Office.context.ui.displayDialogAsync(C.HOST_URL + "/auth.html" + (hint ? "?hint=" + encodeURIComponent(hint) : ""), { height: 60, width: 30, promptBeforeOpen: false }, function (res) {
        diag("dialog sonucu: " + res.status + (res.error ? " " + res.error.code + " " + res.error.message : ""));
        if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error("Giriş penceresi açılamadı (" + res.error.code + "): " + res.error.message + " — alttaki 'panelin içinde giriş yap' bağlantısını kullanın.")); return; }
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
    if (haveToken() || cachedToken()) return Promise.resolve();
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
  /* Süzgece giren METİN değeri adrese kodlanarak girer. 30.09.2026 canlı ölçüm: Dataverse
     ham '+'yı boşluk okuyor (contains(subject,'+') boşluklu konuları döndürdü), ham '&' 400
     veriyor, '#' adresi kesiyor. Gmail message-id'lerinde '+' ve '=' olağan — kodlanmazsa
     "bu e-posta zaten kayıtlı" denetimi ıskalıyor ve aynı e-posta ikinci kez yazılıyordu. */
  function lit(v) { return encodeURIComponent(q(v)); }

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
  // Tarihler Türk biçiminde metin: "gg.aa.yyyy" ve "gg.aa.yyyy ss:dd" (ayraç . / - olabilir).
  var p2 = function (n) { return (n < 10 ? "0" : "") + n; };
  function fmtDate(d) { return p2(d.getDate()) + "." + p2(d.getMonth() + 1) + "." + d.getFullYear(); }
  function localDT(d) { return fmtDate(d) + " " + p2(d.getHours()) + ":" + p2(d.getMinutes()); }
  function parseTr(v, defHour) {
    var m = /^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{4})(?:\s+(\d{1,2})[:.](\d{2}))?\s*$/.exec(v || "");
    if (!m) return null;
    var hh = m[4] != null ? +m[4] : (defHour || 0), mm = m[5] != null ? +m[5] : 0;
    var d = new Date(+m[3], +m[2] - 1, +m[1], hh, mm, 0, 0);
    return isNaN(d.getTime()) || d.getDate() !== +m[1] ? null : d;
  }
  function isoFromLocal(v) { var d = parseTr(v); if (v && !d) throw new Error("Tarih biçimi: gg.aa.yyyy ss:dd"); return d ? d.toISOString() : null; }
  function isoFromDate(v, hour) { var d = parseTr(v, hour ? +hour.split(":")[0] : 9); if (v && !d) throw new Error("Tarih biçimi: gg.aa.yyyy"); return d ? d.toISOString() : null; }

  /* ---------- takvim seçici: metin alanı Türk biçiminde kalır, 📅 tarayıcının kendi seçicisini açar ---------- */
  function bindPickers() {
    document.querySelectorAll(".dp .cal").forEach(function (btn) {
      var txt = $(btn.dataset.for), nat = $(btn.dataset.for + "_n");
      var withTime = nat.type === "datetime-local";
      btn.addEventListener("click", function () {
        var d = parseTr(txt.value, 9) || new Date();
        nat.value = d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate()) + (withTime ? "T" + p2(d.getHours()) + ":" + p2(d.getMinutes()) : "");
        if (typeof nat.showPicker === "function") { try { nat.showPicker(); return; } catch (e) { /* aşağı düş */ } }
        nat.style.pointerEvents = "auto"; nat.style.opacity = "1"; nat.style.width = "auto"; nat.style.height = "auto"; nat.focus(); nat.click();
      });
      nat.addEventListener("change", function () {
        if (!nat.value) return;
        var d = new Date(nat.value);
        txt.value = withTime ? localDT(d) : fmtDate(d);
        nat.style.pointerEvents = ""; nat.style.opacity = ""; nat.style.width = ""; nat.style.height = "";
      });
    });
  }

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
    var m = state.mail; if (!m || !m.fromEmail) return checkExistingEmail();
    // Sabitlenmiş panelde e-posta değişirse eski e-postanın eşlemesi yenisinin üstüne yazılmasın.
    var cur = function () { return state.mail === m; };
    var email = m.fromEmail.toLowerCase(), domain = email.split("@")[1] || "";
    var hint = [];
    return api("GET", "contacts?$select=fullname,contactid,emailaddress1,jobtitle,_parentcustomerid_value&$expand=parentcustomerid_account($select=name,accountid)&$filter=statecode eq 0 and (emailaddress1 eq '" + q(email) + "' or emailaddress2 eq '" + q(email) + "')&$top=3")
      .then(function (r) {
        if (!cur()) return null;
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
        if (!cur()) return;
        if (r && r.value.length === 1) { setAccount({ id: r.value[0].accountid, label: r.value[0].name }); hint.push("Firma alan adından bulundu: " + r.value[0].name + "."); }
        else if (r && r.value.length > 1) hint.push("Alan adı " + r.value.length + " firmayla eşleşti, elle seçin.");
        if (!state.account && !state.contacts.length) hint.push("Gönderen CRM'de bulunamadı; firmayı elle seçin.");
        $("matchHint").textContent = hint.join(" ");
      })
      .then(checkExistingEmail);
  }

  function checkExistingEmail() {
    var m = state.mail, mid = m && m.messageId; if (!mid) return Promise.resolve();
    return api("GET", "emails?$select=activityid,subject&$filter=messageid eq '" + lit(mid) + "'&$top=1").then(function (r) {
      if (r.value.length && state.mail === m) {
        m.existingEmailId = r.value[0].activityid;
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
      api("GET", "WhoAmI"),
      api("GET", "systemusers?$select=fullname,systemuserid&$filter=isdisabled eq false and islicensed eq true and accessmode eq 0 and applicationid eq null&$orderby=fullname")
    ]).then(function (res) {
      var lbl = function (o) { return (o.Label.UserLocalizedLabel || {}).Label || String(o.Value); };
      fillSelect($("tType"), res[0].OptionSet.Options.map(function (o) { return { value: o.Value, label: lbl(o) }; }), true);
      fillSelect($("eFair"), res[1].value.map(function (f) { return { value: f.zeno_exhibitionid, label: f.zeno_name }; }), true);
      fillSelect($("aSeason"), res[2].OptionSet.Options.map(function (o) { return { value: o.Value, label: lbl(o) }; }), true);
      state.userId = res[3].UserId;
      var users = res[4].value.map(function (u) { return { value: u.systemuserid, label: u.fullname }; });
      fillSelect($("tRemindTo"), users, true); fillSelect($("tRemindTo2"), users, true); fillSelect($("tOwner"), users, false);
      if (users.some(function (u) { return u.value === state.userId; })) { $("tRemindTo").value = state.userId; $("tOwner").value = state.userId; }
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
      messageId: it.internetMessageId || "", itemId: it.itemId || "", text: "", html: "",
      // Ek/gömülü görsel içeriği BU öğeden okunur: sabitlenmiş panelde kayıt yoldayken
      // Office.context.mailbox.item çoktan yeni e-postayı gösteriyor olabilir.
      item: it,
      // yalnız gerçek dosya ekleri (gövdeye gömülü görseller ve eklenmiş öğeler hariç)
      attachments: (it.attachments || []).filter(function (a) { return !a.isInline && a.attachmentType === "file"; })
        .map(function (a) { return { id: a.id, name: a.name, size: a.size, type: a.contentType || "" }; }),
      // gövdeye gömülü görseller (imza, ekran görüntüsü): cid: referansı CRM'de çözülmez, base64 olarak gövdeye gömülür
      inline: (it.attachments || []).filter(function (a) { return a.isInline && /^image\//.test(a.contentType || ""); })
        .map(function (a) { return { id: a.id, name: a.name, size: a.size, type: a.contentType }; })
    };
    if (m.attachments.length) {
      $("fSaveAtt").hidden = false;
      $("saveAttLbl").textContent = "Ekleri de kaydet (" + m.attachments.length + "): " + m.attachments.map(function (a) { return a.name; }).join(", ");
    }
    state.mail = m;
    $("mSubject").textContent = m.subject || "(konu yok)";
    $("mFrom").textContent = m.fromName ? m.fromName + " <" + m.fromEmail + ">" : m.fromEmail;
    $("mDate").textContent = m.date.toLocaleString("tr-TR");
    $("subject").value = m.subject;
    // Gövde gelene kadar kayıt KAPALI: açık kalsaydı yeni konu + ÖNCEKİ e-postanın
    // açıklaması ve boş HTML gövdesiyle kayıt açılabiliyordu (sabitlenmiş panel).
    state.loading = true; syncSaveBtn();
    function bitti() { if (state.mail === m) { state.loading = false; syncSaveBtn(); } }
    return new Promise(function (resolve) {
      it.body.getAsync(Office.CoercionType.Text, function (r) {
        m.text = r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : "";
        if (state.mail === m) $("desc").value = buildDescription(m);   // geç gelen eski e-posta gövdesi yenisini ezmesin
        it.body.getAsync(Office.CoercionType.Html, function (r2) {
          m.html = r2.status === Office.AsyncResultStatus.Succeeded ? (r2.value || "") : "";
          bitti();
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
    var ro = isoFromLocal($("tRemindOn").value);
    if (ro) {
      b.zeno_remindon = ro;
      if (!$("tRemindTo").value) throw new Error("Hatırlatma için kime hatırlatılacağını seçin.");
    }
    if ($("tRemindTo").value) b["zeno_RemindToUser_Task@odata.bind"] = "/systemusers(" + $("tRemindTo").value + ")";
    if ($("tRemindTo2").value) b["zeno_remindtouser2_Task@odata.bind"] = "/systemusers(" + $("tRemindTo2").value + ")";
    b.zeno_remindalso = $("tRemindAlso").checked;
    if ($("tOwner").value) b["ownerid@odata.bind"] = "/systemusers(" + $("tOwner").value + ")";
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

  function saveEmailActivity(force) {
    var m = state.mail;
    if ((!force && !$("saveMail").checked) || m.existingEmailId) return Promise.resolve(null);
    var parties = [];
    var sender = state.contacts.filter(function (c) { return (c.email || "").toLowerCase() === m.fromEmail.toLowerCase(); })[0];
    if (sender) parties.push({ "partyid_contact@odata.bind": "/contacts(" + sender.id + ")", participationtypemask: 1 });
    else parties.push({ addressused: m.fromEmail, participationtypemask: 1 });
    m.to.forEach(function (r) { parties.push({ addressused: r.email, participationtypemask: 2 }); });
    m.cc.forEach(function (r) { parties.push({ addressused: r.email, participationtypemask: 3 }); });
    return embedInlineImages(m).then(function (html) {
    if (html.length > 2000000) html = html.slice(0, 2000000) + "<p>[… kısaltıldı]</p>";
    var b = {
      subject: m.subject || "(konu yok)", description: html, directioncode: false, messageid: m.messageId || null,
      sender: m.fromEmail, torecipients: m.to.map(function (r) { return r.email; }).join(";"),
      actualstart: m.date.toISOString(), actualend: m.date.toISOString(), email_activity_parties: parties
    };
    var r = regardingBind("email"); if (r) b[r.key] = r.val;
    return api("POST", "emails", b)
      .then(function (res) { return saveAttachments(res.id).then(function () { return res; }); })
      .then(function (res) {
        // Gelen posta olarak kapat (statecode 1 / statuscode 4 = Alındı). Onay eklentisi yalnız zeno_approvalstatus=1 olanı durdurur.
        return api("PATCH", "emails(" + res.id + ")", { statecode: 1, statuscode: 4 }).then(function () { return res.id; }, function () { return res.id; });
      });
    });
  }

  var INLINE_LIMIT = 3 * 1024 * 1024;
  /* <img src="cid:…"> (ya da http/data olmayan her src) sırayla gömülü görsellerle eşlenir —
     Office.js okuma modunda content-id vermediği için eşleme SIRAYA dayanır; artan görseller sona eklenir. */
  function embedInlineImages(m) {
    var html = m.html || "", list = m.inline || [];
    if (!list.length) return Promise.resolve(html);
    return Promise.all(list.map(function (a) {
      if (a.size > INLINE_LIMIT) return Promise.resolve(null);
      return attachmentContent(a.id, m.item).then(function (b64) { return "data:" + a.type + ";base64," + b64; }, function () { return null; });
    })).then(function (uris) {
      var i = 0, used = [];
      html = html.replace(/(<img\b[^>]*?\bsrc=)(["'])(?!https?:|data:)[^"']*\2/gi, function (all, pre, qt) {
        while (i < uris.length && !uris[i]) i++;
        if (i >= uris.length) return all;
        used.push(i); return pre + qt + uris[i++] + qt;
      });
      var rest = uris.filter(function (u, k) { return u && used.indexOf(k) < 0; });
      if (rest.length) html += "<p>Gömülü görseller:</p>" + rest.map(function (u) { return '<p><img src="' + u + '" style="max-width:100%"></p>'; }).join("");
      return html;
    });
  }

  var ATT_LIMIT = 40 * 1024 * 1024; // Dataverse maxuploadfilesize ~44 MB (canlı ölçüm 14.09)
  function attachmentContent(id, item) {
    return new Promise(function (resolve, reject) {
      (item || Office.context.mailbox.item).getAttachmentContentAsync(id, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error(r.error && r.error.message || "ek okunamadı")); return; }
        if (r.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) { reject(new Error("desteklenmeyen ek biçimi: " + r.value.format)); return; }
        resolve(r.value.content);
      });
    });
  }
  // Ekler sırayla yazılır; biri düşerse diğerleri devam eder, sonuç mesajda özetlenir.
  function saveAttachments(emailId) {
    var list = state.mail.attachments || [], item = state.mail.item;
    if (!$("saveAtt").checked || !list.length) return Promise.resolve();
    state.attResult = { ok: 0, fail: [] };
    return list.reduce(function (chain, a) {
      return chain.then(function () {
        if (a.size > ATT_LIMIT) { state.attResult.fail.push(a.name + " (çok büyük)"); return; }
        return attachmentContent(a.id, item).then(function (b64) {
          return api("POST", "activitymimeattachments", {
            "objectid_email@odata.bind": "/emails(" + emailId + ")", objecttypecode: "email",
            filename: a.name, subject: a.name, mimetype: a.type || "application/octet-stream", body: b64
          });
        }).then(function () { state.attResult.ok++; }, function (e) { state.attResult.fail.push(a.name + " (" + e.message + ")"); });
      });
    }, Promise.resolve());
  }
  function attSummary() {
    var r = state.attResult; if (!r) return "";
    var s = r.ok ? " " + r.ok + " ek kaydedildi." : "";
    if (r.fail.length) s += " Kaydedilemeyen ek: " + r.fail.join("; ") + ".";
    return s;
  }

  /* MÜKERRER KAYIT (30.09.2026, Emel Hn: "kapattığım görev hâlâ listede"). Canlıda 5 görev
     çifti: ikinci kopya 5-15 sn sonra, e-postasız — başarıdan sonra düğme yeniden açılıyordu,
     kullanıcı ikinci kez bastı; birini kapatınca ikizi günlük listede kaldı. Üç kat:
     1) Başarıdan sonra düğme KİLİTLİ kalır, yerinde "Oluşturuldu — CRM'de aç" bağlantısı
        durur. Kilit yalnız başka sekme ya da başka e-posta yüklenince kalkar.
     2) Kaydedilen e-postanın kimliği e-postanın durumuna yazılır: tekrar kayıt aynı e-postayı
        ikinci kez POST etmez, yeni görevi var olan e-postaya bağlar.
     3) Görev POST'undan önce sunucuya sorulur (findOpenTwins): bu e-postaya bağlı açık görev
        ya da aynı kullanıcının son 10 dakikada aynı konuyla açtığı görev varsa panelde onay
        istenir. Office eklentisinde confirm()/alert() çalışmaz; soru panelin içinde.
     Sorgu düşerse kayıt ENGELLENMEZ — denetim yardımcıdır, kullanıcının işini durdurmaz.
     Durum: busy = istek yolda, asking = onay bekleniyor, done = kilit {tab, etn, id, text}. */
  /* Yayın geçişi: GitHub Pages ?v= damgasını yok sayar; önbellekteki ESKİ taskpane.html yeni
     betikle açılabilir. O sayfada btnOpen yok — bağlantı burada eklenir, düğme ölü kalmaz. */
  function openLink() {
    var a = $("btnOpen");
    if (!a) {
      a = document.createElement("a"); a.id = "btnOpen"; a.className = "btn done"; a.target = "_blank"; a.rel = "noopener"; a.hidden = true;
      var b = $("btnSave"); if (b && b.parentNode) b.parentNode.appendChild(a);
    }
    return a;
  }
  /* Bilerek ikinci kayıt (ör. aynı e-postadan başka birine görev): kilit tek tıkla kalkar.
     Kaza kopyasını yine findOpenTwins sorusu yakalar, kilit zayıflamaz. HTML'de değil,
     burada eklenir — önbellekteki eski taskpane.html'de de çalışsın. */
  var ANOTHER_TEXT = { task: "Bu e-postadan başka görev aç", appt: "Bu e-postadan başka toplantı aç", expo: "Bu e-postadan başka fuar notu aç" };
  function anotherBtn() {
    var b = $("btnAnother");
    if (!b) {
      b = document.createElement("button"); b.id = "btnAnother"; b.type = "button"; b.className = "btn sec"; b.hidden = true;
      var o = openLink(); if (o && o.parentNode) o.parentNode.appendChild(b);
    }
    if (!b.liyBound) {
      b.liyBound = true;
      b.addEventListener("click", function () {
        if (!state.done || state.done.tab !== state.tab || state.busy) return;
        state.done = null; showMsg("", ""); syncSaveBtn();
      });
    }
    return b;
  }
  function syncSaveBtn() {
    var btn = $("btnSave"), open = openLink(), again = anotherBtn();
    var d = state.done && state.done.tab === state.tab ? state.done : null;
    btn.hidden = !!d;
    btn.disabled = !!(d || state.busy || state.asking || state.loading || !state.mail);
    open.hidden = !d;
    again.hidden = !(d && ANOTHER_TEXT[d.tab]);
    if (!again.hidden) again.textContent = ANOTHER_TEXT[d.tab];
    if (d) {
      open.textContent = d.text;
      if (d.id) open.href = recordUrl(d.etn, d.id); else open.removeAttribute("href");
    }
  }
  // Uçuştaki kayıt bitti: kilitle ya da serbest bırak; bu sırada e-posta değiştiyse şimdi yükle.
  function endSave(tab, done) {
    state.busy = false;
    state.done = done ? { tab: tab, etn: done.etn, id: done.id, text: done.text || "Oluşturuldu — CRM'de aç" } : null;
    if (state.reloadPending) { state.reloadPending = false; reloadItem(true); return; }
    syncSaveBtn();
  }

  function onSave(ev, confirmed) {
    if (state.done && state.done.tab === state.tab) return;   // kilitli: ne görev ne e-posta yazılır
    if (state.busy || state.asking || state.loading) return;   // loading: gövde henüz okunmadı
    if (!state.mail) { showMsg("err", "E-posta okunamadı; paneli kapatıp yeniden açın."); return; }
    var tab = state.tab, m = state.mail;
    state.busy = true; syncSaveBtn(); showMsg("", "");
    if (tab === "mail") {
      if (!state.account && !state.contacts.length) { showMsg("err", "Bağlanacak firma ya da kişi seçin."); endSave(tab); return; }
      if (m.existingEmailId) { showMsg("ok", "Bu e-posta zaten CRM'de kayıtlı.", { href: recordUrl("email", m.existingEmailId), text: "CRM'de aç" }); endSave(tab); return; }
      saveEmailActivity(true)
        .then(function (id) {
          m.existingEmailId = id;
          showMsg("ok", "E-posta CRM'e kaydedildi ve bağlandı." + attSummary(), { href: recordUrl("email", id), text: "CRM'de aç" });
          endSave(tab, { etn: "email", id: id, text: "Kaydedildi — CRM'de aç" });
        }, function (e) { showMsg("err", "E-posta kaydedilemedi: " + e.message); endSave(tab); });
      return;
    }
    var spec;
    try {
      if (!$("subject").value.trim()) throw new Error("Konu boş olamaz.");
      spec = tab === "task" ? buildTask() : tab === "appt" ? buildAppt() : buildExpo();
    } catch (e) { showMsg("err", e.message); endSave(tab); return; }
    var guard = spec.etn === "task" && !confirmed ? findOpenTwins(spec.body.subject) : Promise.resolve([]);
    guard.then(function (twins) {
      if (state.reloadPending) { endSave(tab); return; }        // bu arada başka e-postaya geçildi
      if (twins.length) { askDuplicate(twins); return; }
      return createRecord(spec, m, tab);
    }).catch(function (e) { showMsg("err", "Kayıt oluşturulamadı: " + e.message); endSave(tab); });
  }

  function createRecord(spec, m, tab) {
    var created = null;
    return api("POST", spec.set, spec.body)
      .then(function (r) { created = r.id; return saveEmailActivity(); })
      .then(function (emailId) {
        if (emailId) m.existingEmailId = emailId;   // 2) tekrar kayıtta e-posta ikinci kez POST edilmez
        return linkTaskToEmail(spec, created, emailId).then(function (warn) { return { emailId: emailId, warn: warn }; });
      })
      .then(function (r) {
        var link = { href: recordUrl(spec.etn, created), text: "CRM'de aç" };
        if (r.warn) showMsg("err", r.warn, link);
        else showMsg("ok", spec.label + " oluşturuldu." + (r.emailId ? " E-posta da kaydedildi." + attSummary() : ""), link);
        endSave(tab, { etn: spec.etn, id: created });
      })
      .catch(function (e) {
        if (created) {
          // Kayıt açıldı: yeniden denemek ikinci kopya açar, düğme kilitli kalır.
          showMsg("err", spec.label + " oluşturuldu ama e-posta kaydedilemedi: " + e.message + " — tekrar basmayın; e-postayı 'Yalnız E-posta' sekmesinden kaydedebilirsiniz.", { href: recordUrl(spec.etn, created), text: "CRM'de aç" });
          endSave(tab, { etn: spec.etn, id: created });
        } else { showMsg("err", "Kayıt oluşturulamadı: " + e.message); endSave(tab); }
      });
  }

  var TWIN_MINUTES = 10;
  function rowsOf(r) { return (r && r.value) || []; }
  // Her ayak kendi hatasını yutar: denetim düşerse boş liste döner, kayıt engellenmez.
  function findOpenTwins(subject) {
    var byMail = Promise.resolve()
      .then(function () { return state.mail.existingEmailId ? null : checkExistingEmail(); })
      .then(function () {
        var eid = state.mail.existingEmailId;
        if (!eid) return [];
        return api("GET", "tasks?$select=activityid,subject&$filter=_zeno_emailid_value eq " + eid + " and statecode eq 0&$orderby=createdon desc&$top=5").then(rowsOf)
          .then(function (rows) { return rows.map(function (t) { return { activityid: t.activityid, subject: t.subject, via: "mail" }; }); });
      })
      .catch(function (e) { diag("mükerrer denetimi (e-posta) düştü: " + e.message); return []; });
    var bySubject = Promise.resolve()
      .then(function () {
        var since = new Date(Date.now() - TWIN_MINUTES * 60000).toISOString();
        var f = "subject eq '" + lit(subject) + "' and createdon ge " + since;
        if (state.userId) f += " and _createdby_value eq " + state.userId;
        return api("GET", "tasks?$select=activityid,subject&$filter=" + f + "&$orderby=createdon desc&$top=5").then(rowsOf);
      })
      .catch(function (e) { diag("mükerrer denetimi (konu) düştü: " + e.message); return []; });
    return Promise.all([byMail, bySubject]).then(function (r) {
      var seen = {}, out = [];
      r[0].concat(r[1]).forEach(function (t) {
        if (t && t.activityid && !seen[t.activityid]) { seen[t.activityid] = true; out.push({ id: t.activityid, subject: t.subject || "", via: t.via || "subject" }); }
      });
      return out;
    });
  }

  function askDuplicate(twins) {
    state.busy = false; state.asking = true; syncSaveBtn();
    // "err warn": eski önbellekli sayfada .msg.warn kuralı yok, .msg.err kutuyu yine görünür kılar.
    var el = $("msg"); el.className = "msg err warn"; el.textContent = "";
    // Konu ayağı BAŞKA bir e-postadan açılmış görevi de bulabilir (aynı konulu yazışma):
    // "bu e-postadan" demek o zaman yanlış olurdu.
    var fromMail = twins.some(function (t) { return t.via === "mail"; });
    el.appendChild(document.createTextNode(fromMail
      ? "Bu e-postadan açılmış görev var — yine de yeni görev açılsın mı? "
      : "Son " + TWIN_MINUTES + " dakikada aynı konulu bir görev açtınız — yine de yeni görev açılsın mı? "));
    var a = document.createElement("a"); a.href = recordUrl("task", twins[0].id); a.target = "_blank"; a.rel = "noopener";
    a.textContent = (twins[0].subject || "Görev") + " — CRM'de aç"; el.appendChild(a);
    if (twins.length > 1) el.appendChild(document.createTextNode(" (+" + (twins.length - 1) + " görev daha)"));
    var row = document.createElement("div"); row.className = "ask";
    var yes = document.createElement("button"); yes.type = "button"; yes.className = "btn"; yes.textContent = "Evet, yeni görev aç";
    var no = document.createElement("button"); no.type = "button"; no.className = "btn sec"; no.textContent = "Vazgeç";
    yes.addEventListener("click", function () { if (!state.asking) return; state.asking = false; onSave(null, true); });
    no.addEventListener("click", function () { if (!state.asking) return; state.asking = false; showMsg("", ""); syncSaveBtn(); });
    row.appendChild(yes); row.appendChild(no); el.appendChild(row);
    try { no.focus(); } catch (e) { /* odak zorunlu değil */ }
  }

  /* GÖREV → KAYNAK E-POSTA (14.09.2026, Furkan: "task üzerinde email lookup'ını da
     koyabilir miyiz?"). Görevde toplantı/telefon kaynak lookup'ları gibi zeno_emailid
     (Kaynak E-posta) var; bind adı metadata'dan: zeno_EmailId_Task. E-posta bu turda
     kaydedildiyse o, zaten CRM'de varsa mevcut kayıt bağlanır. Bağ kurulamazsa görev
     yine de kalır — hata mesajda söylenir, kayıt geri alınmaz. Uyarı METİN olarak döner:
     eskiden burada gösterilip hemen ardından "oluşturuldu" mesajıyla eziliyordu. */
  function linkTaskToEmail(spec, taskId, emailId) {
    var eid = emailId || (state.mail && state.mail.existingEmailId);
    if (spec.etn !== "task" || !taskId || !eid) return Promise.resolve();
    return api("PATCH", "tasks(" + taskId + ")", { "zeno_EmailId_Task@odata.bind": "/emails(" + eid + ")" })
      .then(function () { return ""; }, function (e) { return "Görev oluşturuldu ama e-postaya bağlanamadı: " + e.message; });
  }

  /* ---------- sekmeler ---------- */
  function selectTab(name) {
    if (name !== state.tab) {
      // Kilit yalnız başka sekmeye geçince kalkar (aynı sekmede kutu değişimi kilidi açmaz); bekleyen soru düşer.
      if (state.asking) { state.asking = false; showMsg("", ""); }
      state.done = null;
    }
    state.tab = name;
    document.querySelectorAll(".tabs [role=tab]").forEach(function (b) { b.setAttribute("aria-selected", String(b.dataset.tab === name)); });
    $("pTask").hidden = name !== "task"; $("pAppt").hidden = name !== "appt"; $("pExpo").hidden = name !== "expo"; $("pMail").hidden = name !== "mail";
    var mailOnly = name === "mail";
    $("fSubject").hidden = mailOnly; $("fDesc").hidden = mailOnly; $("fSaveMail").hidden = mailOnly;
    $("fSaveAtt").hidden = !(state.mail && state.mail.attachments.length) || (!mailOnly && !$("saveMail").checked);
    $("btnSave").textContent = mailOnly ? "E-postayı CRM'e bağla" : "CRM'de oluştur";
    syncSaveBtn();
  }

  /* ---------- sabitlenmiş panel: başka e-posta seçilince ---------- */
  /* manifest SupportsPinning=true; ItemChanged dinlenmezse panel eski e-postanın konusu, kişisi
     ve kilidiyle kalıyordu. Kayıt yoldayken gelen değişim, kayıt bitince uygulanır — yoksa
     yolda olan kayıt yeni e-postanın durumuna yazardı. */
  function onItemChanged() {
    if (state.busy) { state.reloadPending = true; return; }
    reloadItem(false);
  }
  function reloadItem(keepMsg) {
    state.done = null; state.asking = false; state.attResult = null;
    if (!keepMsg) showMsg("", "");
    setAccount(null); state.contacts = []; renderContacts();
    $("matchHint").textContent = ""; $("mLinked").textContent = "";
    $("desc").value = "";                          // önceki e-postanın açıklaması yeni konuyla kalmasın
    $("saveMail").checked = true; $("saveAtt").checked = true; $("fSaveAtt").hidden = true;
    if (!Office.context.mailbox.item) {           // Outlook masaüstü: hiçbir öğe seçili değil
      state.mail = null; state.loading = false;
      $("mSubject").textContent = "—"; $("mFrom").textContent = ""; $("mDate").textContent = "";
      $("subject").value = ""; $("desc").value = "";
      syncSaveBtn(); return;
    }
    var p;
    try { p = readMail(); } catch (e) { state.mail = null; state.loading = false; showMsg("err", "E-posta okunamadı: " + e.message); syncSaveBtn(); return; }
    syncSaveBtn();
    p.then(autoMatch).catch(function (e) { showMsg("err", e.message); });
  }
  var itemWatch = false;
  function watchItem() {
    if (itemWatch) return; itemWatch = true;
    try {
      var mb = Office.context.mailbox;
      if (mb && mb.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
        mb.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, function (r) {
          if (r && r.status !== Office.AsyncResultStatus.Succeeded) diag("ItemChanged kaydedilemedi: " + (r.error && r.error.message));
        });
      }
    } catch (e) { diag("ItemChanged: " + e.message); }
  }

  /* ---------- başlat ---------- */
  function boot() {
    document.querySelectorAll(".tabs [role=tab]").forEach(function (b) { b.addEventListener("click", function () { selectTab(b.dataset.tab); }); });
    lookup("accQ", "accList", searchAccounts, setAccount);
    lookup("conQ", "conList", searchContacts, addContact);
    $("btnSave").addEventListener("click", onSave);
    $("saveMail").addEventListener("change", function () { selectTab(state.tab); });
    bindPickers();
    $("btnLogin").addEventListener("click", function () {
      diag("giriş düğmesi tıklandı");
      $("loginMsg").className = "msg ok"; $("loginMsg").textContent = "Giriş penceresi açılıyor…";
      dialogToken().then(start).catch(function (e) { $("loginMsg").className = "msg err"; $("loginMsg").textContent = e.message; });
    });
    // Yedek: Office dialog açılmayan istemcilerde (eski Outlook Mac) giriş panelin kendi penceresinde yapılır.
    $("btnLoginInline").addEventListener("click", function (ev) {
      ev.preventDefault(); diag("panel içi giriş tıklandı");
      try { localStorage.setItem("liy_back", "1"); } catch (e) { diag("localStorage yok: " + e.message); }
      var hint = state.user ? state.user.username : "";
      msal.initialize().then(function () {
        return msal.acquireTokenRedirect(hint ? { scopes: SCOPES, loginHint: hint } : { scopes: SCOPES, prompt: "select_account" });
      }).catch(function (e) { $("loginMsg").className = "msg err"; $("loginMsg").textContent = e.message; });
    });
    try { var le = localStorage.getItem("liy_err"); if (le) { localStorage.removeItem("liy_err"); $("loginMsg").className = "msg err"; $("loginMsg").textContent = le; } } catch (e) { /* yok */ }
    // varsayılan tarihler
    var d = new Date(); d.setDate(d.getDate() + 3); $("tDue").value = fmtDate(d);
    var rd = new Date(d); rd.setHours(8, 30, 0, 0); $("tRemindOn").value = localDT(rd);
    var s = new Date(); s.setDate(s.getDate() + 1); s.setHours(10, 0, 0, 0); $("aStart").value = localDT(s);
    var e = new Date(s.getTime() + 60 * 60000); $("aEnd").value = localDT(e);
    $("eDate").value = fmtDate(new Date());

    if (cachedToken()) { start(); return; }
    silentToken().then(start, function () { $("login").style.display = "block"; });
  }

  function start() {
    $("login").style.display = "none"; $("app").hidden = false;
    watchItem();
    readMail().then(loadLists).then(autoMatch).catch(function (e) { showMsg("err", e.message); });
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) { document.body.textContent = "Bu eklenti yalnız Outlook'ta çalışır."; return; }
    try {
      var d = Office.context.diagnostics || {};
      diag("Office hazır: " + info.host + "/" + info.platform + " v" + (d.version || "?") + " · Mailbox1.8=" + Office.context.requirements.isSetSupported("Mailbox", "1.8") + " · dialog=" + (Office.context.ui && typeof Office.context.ui.displayDialogAsync) + " · msal=" + (typeof window.msal));
    } catch (e) { diag("tanı hatası: " + e.message); }
    boot();
  });
})();
