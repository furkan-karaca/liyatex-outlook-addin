# Liyatex CRM — Outlook eklentisi

Açık e-postadan Liyatex CRM'de **Görev / Toplantı / Fuar Notu** açar, göndereni CRM'deki
kişi/firmayla otomatik eşler ve isteğe bağlı e-postayı `email` aktivitesi olarak firmaya bağlar.

Neden özel eklenti: kullanıcıların posta kutusu CRM kiracısından (liyatexcrm) farklı bir
kiracıda; Microsoft'un "Dynamics 365 App for Outlook"u çapraz kiracıda desteklenmiyor. Bu
eklenti kiracıya bakmaz — kullanıcı panelde **CRM hesabıyla** (`…@liyatexcrm.onmicrosoft.com`)
oturum açar, veri doğrudan tarayıcıdan Dataverse Web API'ye yazılır.

## Dosyalar

| Dosya | İş |
|---|---|
| `manifest.xml` | Outlook'a eklenen şerit düğmesi (Okuma modunda "CRM'e Kaydet") |
| `taskpane.html/.js` | Panel: e-posta özeti, firma/kişi eşleme, üç sekme, kaydet |
| `auth.html` | Office iletişim kutusunda MSAL girişi, jetonu panele iletir |
| `config.js` | CLIENT_ID / HOST_URL / CRM adresi |
| `lib/msal-browser.min.js` | MSAL 3.28.1 (yerel kopya, CDN bağımlılığı yok) |
| `assets/` | Simgeler |

## Kurulum (sırayla)

### 1. Uygulama kaydı — liyatexcrm kiracısı (Entra ID)
`entra.microsoft.com` → App registrations → New:
- Ad: **Liyatex Outlook Eklentisi**, hesap türü: *Single tenant*.
- **Authentication → Add a platform → Single-page application**, Redirect URI:
  `https://<HOST>/auth.html`.
- **API permissions → Add → Dynamics CRM → Delegated → `user_impersonation`** → *Grant admin consent*.
- Overview'daki **Application (client) ID**'yi `config.js › CLIENT_ID`'ye yaz.

Var olan `LiyatexCrmApp` (S2S) kaydını kullanma: o kayıt gizli anahtarlı arka plan kimliği,
tarayıcıda SPA akışıyla karıştırmak yanlış.

### 2. Yayımlama (HTTPS zorunlu)
Klasör tamamen statik. Seçenekler: Azure Storage *static website* (kartela görsellerinin
durduğu hesapta `$web` kapsayıcısı), Azure Static Web Apps, ya da GitHub Pages.
Adres belli olunca:

```bash
./set-host.sh https://liyatex-outlook.example.com
```

(betik `manifest.xml` ve `config.js` içindeki HOST adresini değiştirir).

### 3. Outlook'a ekleme — POSTA kiracısı
Kullanıcıların postası hangi kiracıdaysa orada:
- Tüm kullanıcılara: Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps** → `manifest.xml`.
- Tek kullanıcıya: Outlook → **Eklentiler → Özel eklentilerim → Dosyadan ekle** → `manifest.xml`.

Düğme okuma penceresinin şeridinde "Liyatex CRM › CRM'e Kaydet" olarak çıkar. İlk açılışta
CRM hesabıyla giriş ister; oturum tarayıcıda (localStorage) kalır.

## Yerel deneme (yayımlamadan)
```bash
npx office-addin-dev-certs install       # bir kere; localhost sertifikası
./set-host.sh https://localhost:3000
npx http-server . -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```
Sonra manifest'i "Dosyadan ekle" ile Outlook'a yükle. Uygulama kaydında `https://localhost:3000/auth.html`
redirect URI olarak da eklenmeli.

## Ne yazıyor (doğrulandı, 11.09.2026 — S2S kimliğiyle canlıda oluşturulup silindi)
- **Görev** `tasks`: subject, description, prioritycode, `zeno_gorevturu`, `zeno_kaynakturu=Mail`, scheduledend, `regardingobjectid_account_task`.
- **Toplantı** `appointments`: başlangıç/bitiş, location, `zeno_sezon`, `zeno_TakenById_Appointment`=kullanıcı, `regardingobjectid_account_appointment`, `appointment_activity_parties` (kişiler, mask 5), `zeno_customerattendees` metni.
- **Fuar Notu** `zeno_exhibitionnotes`: `zeno_exhibitionid_zeno_exhibitionnote`, `zeno_interactiontypecode`, `zeno_visitorid_zeno_exhibitionnote`=kullanıcı, `regardingobjectid_account_zeno_exhibitionnote`, `zeno_exhibitionnote_activity_parties` (İlgili Kişiler), `zeno_ContactId_zeno_exhibitionnote`.
- **E-posta** `emails`: messageid (mükerrer kontrolü bununla), HTML gövde, from/to/cc party'leri, `directioncode=false`; ardından `statecode 1 / statuscode 4 (Alındı)`. Onay eklentisi (`OutgoingApprovalGuard`) yalnız `zeno_approvalstatus=1` olan postayı durdurur, gelen posta etkilenmez.

Firma yoksa ilk seçilen kişi "İlgili" olur; ikisi de yoksa kayıt bağsız açılır (panel uyarır).

## Bilinen sınırlar
- office.js Microsoft CDN'inden yüklenir; Microsoft SRI/yerel kopyayı desteklemez.
- Sunucu tarafı eşitleme yok: CRM'de açılan toplantı Outlook takvimine **düşmez** (ve tersi). İstenirse çapraz kiracı server-side sync ayrıca kurulur.
- Yeni Outlook (Windows) ve Outlook Web'de çalışır; klasik Outlook 2016+ da manifest'i kabul eder. Mobilde okuma penceresi düğmesi tanımlı.
