# Çocuk Kulübü Uygulaması · Kurulum (bir kez)

Veriler **alibaklatazof@gmail.com** hesabının Google Drive'ındaki bir e-tabloda durur.

1. Tarayıcıda **alibaklatazof@gmail.com** ile Google Drive'a girin → **Yeni → Google E-Tablolar**. Adını "Çocuk Kulübü Uygulaması · Veriler" yapın.
2. Tabloda **Uzantılar → Apps Script**. Açılan sayfadaki kodun hepsini silin, `apps-script/Kod.gs` dosyasının içeriğini yapıştırın, 💾 Kaydet.
3. Soldaki **Hizmetler (+)** → listeden **Drive API** → sürüm **v3** → **Ekle**. (Dekont okuma için gerekli.)
4. Üstteki işlev listesinden **kurulum** seçin → **Çalıştır**. İzin penceresinde hesabınızı seçin → **Gelişmiş → (güvenli değil) sayfasına git → İzin ver**. Bu uyarı, betiği sizin yazdığınız için çıkar.
   Tabloda Ayarlar, Okullar, Ogrenciler, Dekontlar… sayfaları oluşur. **Ayarlar** sayfasındaki **B2** hücresi okullara vereceğiniz **kayıt kodudur**; istediğiniz zaman değiştirebilirsiniz.
5. Sağ üstte **Dağıt → Yeni dağıtım** → tür: **Web uygulaması** → Yürüten: **Ben**, Erişimi olan: **Herkes** → **Dağıt**. Çıkan **Web uygulaması URL'sini** kopyalayın.
6. Bu URL `config.js` içindeki `API_URL` değerine yazılır ve site güncellenir.

Kodu sonradan değiştirince: **Dağıt → Dağıtımları yönet → ✎ → Sürüm: Yeni sürüm → Dağıt** (URL değişmez).
