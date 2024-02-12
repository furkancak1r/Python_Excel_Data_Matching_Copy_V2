# Bu Python kodu, bir Excel dosyasının 'Sayfa2'sindeki C sütunundaki her hücrenin içeriğini günceller. 
# Her hücredeki değeri kontrol eder ve eğer '-' karakteri içeriyorsa, bu karakterin ilk görünümünden sonraki kısmı alarak hücreyi bu yeni değerle günceller.
# Örneğin, "1-xxxxxxxx" değeri "xxxxxxxx" olarak güncellenir. Yapılan tüm değişiklikler kaydedilir.

import openpyxl

def update_column(file_path):
    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(file_path)

    # Sayfa1'i al
    sheet = workbook['Sayfa2']

    # C sütunundaki her hücre için
    for cell in sheet['C']:
        if cell.value and '-' in cell.value:
            # İlk '-' karakterinden sonrasını al
            updated_value = cell.value.split('-', 1)[1]
            # Güncellenmiş değeri aynı hücreye yaz
            cell.value = updated_value

    # Değişiklikleri kaydet
    workbook.save(file_path)

# Script'i çalıştırmak için:
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\SSH\\SSH Exceller\\ŞOK\\1.xlsx"
update_column(file_path)
