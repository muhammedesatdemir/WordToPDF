import os              #dosya yolu işlemleri için kullanılacak kütüphane
import comtypes.client #office uygulamalarıyla etkileşim kurmak için COM nesneleri oluşturan kütüphane


while True:
#kullanıcı hatalı dosya yolu girdiğinde program tekrar dosya yolunu girmesine izin verecek bir döngü sağlar
    word_dosyasi = input("Lütfen Word dosyasının tam yolunu giriniz\nÖrnek -> C:\\Users\\Monster\\Desktop\\Belgeler\\file.docx\n").strip()
    #kullanıcıdan word dosyasının yolu alınır(strip ile baştaki ve sondaki boşluklar atılır)

    word_dosyasi = word_dosyasi.replace("\\", "/") #kullanıcıya '\' ile verilmesi gereken bilgi kodda hataya sebebiyet veriyordu(\t,\n gibi)
    #ters eğik çizgileri düzeltildi

    if not os.path.isfile(word_dosyasi):  #dosyanın olup olmadığını kontrol eder
        print("Belirtilen dosya bulunamadı.")
        continue

    if not word_dosyasi.endswith(".docx"): #dosyanın uzantısının .docx olup olmadığını kontrol eder
        print("Lütfen bir .docx dosyası seçiniz.")
        continue

    pdf_dosyasi = os.path.splitext(word_dosyasi)[0] + ".pdf"
    #dosyanın adını ve uzantısını ayırır.örnek:file.docx → ("file", ".docx")
    #iki indekse ayrılan ifadenin [0] yani ilk indexinden dosyanın adı alınıp sonuna ".pdf" ekleyerek pdf dosyası adı oluşturulur(file.pdf)
        
    word = comtypes.client.CreateObject("Word.Application")
    #üzerinde işlem yapılacak word uygulaması başlatılır
    
    try:
    #hata olasılığı bulunan kodları denemek için bir try bloğu başlatır,hata oluşursa except bloğuna geçilir

        dokuman_yolu = os.path.abspath(word_dosyasi)
        pdf_yolu = os.path.abspath(pdf_dosyasi)
        #dosya yolları oluşturulur
        
        pdf_format = 17
        #dosyanın pdf olarak kaydedilmesini sağlar
        
        dosya = word.Documents.Open(dokuman_yolu)
        #word dosyasını açar
        
        dosya.SaveAs2(pdf_yolu, FileFormat=pdf_format)
        #pdf olarak kaydeder
        
        dosya.Close()
        #word dosyası kapatılır
        
        print(f"PDF başarıyla oluşturuldu: {pdf_yolu}")
        break
        
    except Exception as e:
        print(f"Hata oluştu: {e}") #oluşan hatanın detaylarını e değişkeninde tutup hata mesajını kullanıcıya gösterir
        break

    finally:
        word.Quit()
    #word uygulaması kapatılır"