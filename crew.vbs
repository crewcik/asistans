MsgBox "Merhaba, ben asistanin."
MsgBox "Seni biraz tanimak istiyorum :)"

adin = InputBox("Adiniz nedir?", "Crew - Asistan's")
MsgBox "Merhaba " & adin & "! Bilgisayarina yardimci olmam icin masaustu yolunu yapistirman gerekecek."
yol = InputBox("Masaüstü Dosya Yolunu Yapistir:", "Crew - Windows Güvenlik Uyarisi")

islem = InputBox("Islem Belirtin: (Dosya Sorgu / Dosya Olustur / Dosya Sil)", "Crew - Soru")

If LCase(islem) = "dosya sorgu" Then
    dosya_adi = InputBox("Dosya adi nedir?", "Crew - Dosya Sorgulama")
    If dosyaVarMi(dosya_adi) Then
        MsgBox "Dosya mevcut."
    Else
        MsgBox "Dosya mevcut degil."
    End If
elseif LCase(islem) = "dosya olustur" Then    
    dosya_adii = InputBox("Dosya adi ne olsun?", "Crew - Dosya Oluşturucu")
    dosya_uzanti = InputBox("Dosya uzantisi ne olsun? (Örn: .html, .js, .txt)", "Crew - Dosya Uzantisi")
    dosyaOlustur(dosya_adii & dosya_uzanti)
elseif LCase(islem) = "dosya sil" Then
    dosya_adii = InputBox("Dosya adi nedir?", "Crew - Dosya Silme")
    dosyaSil(dosya_adii)
else 
    MsgBox "Geçersiz işlem belirttiniz."
End if

Function dosyaVarMi(dosyaAdi)
    Set dosyaSistem = CreateObject("Scripting.FileSystemObject")
    dosyaYolu = yol&dosyaAdi
    If dosyaSistem.FileExists(dosyaYolu) Then
        dosyaVarMi = True
    Else
        dosyaVarMi = False
    End If
End Function

Sub dosyaOlustur(dosyaAdi)
    Set dosyaSistem = CreateObject("Scripting.FileSystemObject")
    dosyaYolu = yol&dosyaAdi
    Set dosya = dosyaSistem.CreateTextFile(dosyaYolu)
    MsgBox "Dosya olusturuldu: " & dosyaAdi
End Sub

Sub dosyaSil(dosyaAdi)
    Set dosyaSistem = CreateObject("Scripting.FileSystemObject")
    dosyaYolu = yol&dosyaAdi
    If dosyaSistem.FileExists(dosyaYolu) Then
        dosyaSistem.DeleteFile(dosyaYolu)
        MsgBox "Dosya silindi: " & dosyaAdi
    Else
        MsgBox "Dosya mevcut degil: " & dosyaAdi
    End If
End Sub
