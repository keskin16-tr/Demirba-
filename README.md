scripts

//nlabel.vbs

' start.vbs
' nlabel-appw.exe dosyasını komut penceresi göstermeden çalıştırır.

Set shell = CreateObject("WScript.Shell")

' 1. Uygulamanın çalışacağı klasörü alın
appPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))

' 2. nlabel-appw.exe'yi arka planda (0) çalıştırın
' NOT: App adını kontrol edin, eğer nlabel-app.exe ise burayı nlabel-app.exe olarak değiştirin.
shell.Run Chr(34) & appPath & "nlabel-appw.exe" & Chr(34), 0, false

' 3. Tarayıcıyı 2 saniye sonra açın (Sunucunun başlaması için zaman tanır)
WScript.Sleep 500 

' 4. Tarayıcıyı http://localhost:3000 adresine yönlendirin
shell.Run "http://localhost:3000/"


//Windows Yazıcı Sürücüsü Ayarı (Kritik Adım)

Tarayıcıdan gönderilen auto komutunun fiziksel olarak cihazda karşılık bulması için Windows denetim masasından Brother yazıcınızın ayarlarına gidin:
1-Yazdırma Tercihleri (Printing Preferences) menüsünü açın.
2-Sayfa Yapısı (Page Setup) sekmesinden "Continuous Tape" (Sürekli Şerit) modunu seçin.
3-Uzunluk (Length) ayarını "Auto" (Otomatik) olarak işaretleyin.
4-Kesim ayarlarından (Kağıt Kesme / Cutter) "Chain Print" (Zincirleme Baskı) veya "Cut at End" (Sonunda Kes) seçeneğini aktif edin.
