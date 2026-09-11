// Liyatex Outlook eklentisi — ortam ayarları.
// CLIENT_ID: liyatexcrm kiracısında açılacak "Liyatex Outlook Eklentisi" uygulama kaydının Application (client) ID'si.
// HOST_URL : eklentinin yayımlandığı HTTPS kök adres (manifest.xml içindeki adresle birebir aynı olmalı).
window.LIY_CFG = {
  CLIENT_ID: "00000000-0000-0000-0000-000000000000",
  TENANT_ID: "e162e6fc-1efd-4892-9039-626a81816b87",
  CRM_URL: "https://liyatexprod.crm4.dynamics.com",
  HOST_URL: "https://furkan-karaca.github.io/liyatex-outlook-addin",
  APP_ID: "f87cad63-3188-f111-ab0e-e4fb1ef65c58", // Fabric uygulaması (kayıt bağlantıları için)
  BODY_LIMIT: 4000 // açıklamaya kopyalanacak azami e-posta metni
};
