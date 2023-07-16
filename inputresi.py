from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from selenium.webdriver.chrome.options import Options
import datetime
import pandas as pd
import re
import csv

download_directory = 'C:\\Users\\Aditya PC\\Downloads\\Surat Jalan'

chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_directory,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': False
})
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"],)
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)

current_time = datetime.datetime.now()
print(f'waktu mulai{current_time}')

driver.get("https://opsbasarta.com/")
find_tipe_akun = \
driver.find_element(By.ID, "tipe_account")
select_tipe_akun = Select(find_tipe_akun)
select_tipe_akun.select_by_value("pengepul")
time.sleep(0.5)

find_nama_akun = \
driver.find_element(By.ID, "nama_account")
select_nama_akun = Select(find_nama_akun)
select_nama_akun.select_by_value("PJG")

driver.find_element("id", "login").send_keys("Rana")
driver.find_element("id", "password").send_keys("54321")
driver.find_element("name", "submit").click()
time.sleep(0.5)

# ===================

# Membuka halaman web
driver.get("https://pool.opsbasarta.com/pengepul/pengepul_home")

# Mengakses elemen input dengan XPath
input = driver.find_element(By. XPATH, '/html/body/div[3]/div/table[1]/tbody/tr/td[1]/form/table/tbody/tr/td[2]/input')

# Mengakses tombol dengan XPath
button_1 = driver.find_element(By.XPATH, '/html/body/div[3]/div/table[1]/tbody/tr/td[1]/form/table/tbody/tr/td[3]/a/button')

# Membaca file CSV
df = pd.read_csv('C:\\Users\Aditya PC\\Downloads\\Surat Jalan\\data.csv')

# Mengambil nilai dari sel A2 hingga baris berikutnya
kolom_resi = df['Resi'][0:].tolist()
resi_pertama = kolom_resi[0]
# Menampilkan nilai yang diambil
print(f"--------{resi_pertama}--------")

input.send_keys(resi_pertama)
button_1.click()

history_barang_tabel = driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/table')
history_barang_row = history_barang_tabel.find_elements(By.TAG_NAME, "tr")
history_barang_row_count = len(history_barang_row)
print(f"Jumlah baris history barang: {history_barang_row_count}")

#++++++LOOP PENCARIAN PERTAMA+++++++++


list_detail_resi = []
with open("C:\\Users\\Aditya PC\\Downloads\\Surat Jalan\\detail-resi.csv", "a", newline="") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Nama Pengirim","Nama Penerima", "Nomor Pengirim", "Nomor Penerima", "Isi", "Ongkir", "Berat", "Tujuan", "Alamat Pengirim","Alamat Penerima"])
    
for e in range (1, history_barang_row_count):
    sel_history_barang = driver.find_element(By.XPATH, f'/html/body/div[3]/div/div[2]/table/tbody/tr[{e}]/td[4]').text
    # print(sel_history_barang)


    # =====ALAMAT=====START
    # Regex Alamat
    pengirim_pattern = r"Alamat:\s(.*?)\s-"
    penerima_pattern = r"Alamat:\s(.*?)\s-.*?Penerima\s\(Alamat:\s(.*?)\s-"

    pengirim_match = re.search(pengirim_pattern, sel_history_barang)
    penerima_match = re.search(penerima_pattern, sel_history_barang)

    alamat_pengirim = None
    alamat_penerima = None
    if pengirim_match and penerima_match:
        alamat_pengirim = pengirim_match.group(1).replace(", ","_")
        alamat_penerima = penerima_match.group(2).replace(", ","_")
        if alamat_pengirim is not None and alamat_penerima is not None:
            fill_alamat_pengirim = alamat_pengirim
            fill_alamat_penerima = alamat_penerima
            
            print("Alamat Pengirim:", fill_alamat_pengirim)
            print("Alamat Penerima:", fill_alamat_penerima)

    # else:
    #     print("Tidak ada alamat pengirim dan penerima")

    # =====ALAMAT=====END


    # =====NAMA=====START

    nama_pengirim = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[2]").text
    nama_penerima = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[3]").text

    if alamat_pengirim is not None and alamat_penerima is not None:
        fill_nama_pengirim = nama_pengirim
        fill_nama_penerima = nama_penerima
        print(f"Nama Pengirim: {fill_nama_pengirim}")
        print(f"Nama Penerima: {fill_nama_penerima}")
    else:
        nama_pengirim = None
        nama_penerima = None

    # =====NAMA=====END


    # =====NO TLP=====START
    # Regex Nomor Telepon
    phone_numbers = re.findall(r'Telp: (\d+)', sel_history_barang)

    # Loop Nomor Telepon
    nomor_pengirim = None
    nomor_penerima = None
    if phone_numbers:
        nomor_pengirim = phone_numbers[0]
        nomor_penerima = phone_numbers[1]
        if nomor_pengirim is not None or nomor_penerima is not None:
            if nomor_pengirim == '-' or nomor_penerima == '-':
                nomor_pengirim = '-'
                nomor_penerima = '-'
            else:
                fill_nomor_pengirim = nomor_pengirim
                fill_nomor_penerima = nomor_penerima
                print(f"Nomor Pengirim: {nomor_pengirim}")
                print(f"Nomor Penerima: {nomor_penerima}")


    # =====NO TLP=====END


    # =====ISI=====START

    isi_barang_pattern = r"Isi:\s(.*?)\),"
    isi_barang_match = re.search(isi_barang_pattern, sel_history_barang)

    data_barang = None
    if isi_barang_match:
        data_barang = isi_barang_match.group(1)
        if data_barang is not None:
            if data_barang == '':
                data_barang = "-"
                fill_data_barang = data_barang
                print(f"Isi Barang: {fill_data_barang}")
            else:
                fill_data_barang = data_barang
                print(f"Isi Barang:{fill_data_barang}")
    # else:
    #     print("Tidak dapat menemukan data barang.")

    # =====ISI=====END


    # =====ONGKIR=====START
    # Regex ongkir
    ongkir_pattern = r"Total:\sRp\.(\d{1,3}(?:,\d{3})*)"
    ongkir_match = re.search(ongkir_pattern, sel_history_barang)

    ongkir = None
    if ongkir_match:
        ongkir_string = ongkir_match.group(1)
        ongkir_string = ongkir_string.replace(",", "")
        ongkir = int(ongkir_string)
        if ongkir is not None:
            fill_ongkir = ongkir
            print("Ongkir:", fill_ongkir)
    # else:
    #     print("Tidak dapat menemukan ongkir.")

    # =====ONGKIR=====END

    # =====BERAT=====START

    kubik_pattern = r'Kubik: (-|\d+\.\d+)'
    kubik_match = re.search(kubik_pattern, str(sel_history_barang))

    kilo_pattern = r'Kilo: (\d+)'
    kilo_match = re.search(kilo_pattern, str(sel_history_barang))

    total_berat_1 = 0
    if kubik_match:
        kubik_value = kubik_match.group(1)
        if kubik_value == '-':
            print(f"Kubik: {kubik_value}")
        else:
            kubik_value = float(kubik_value)
            konversi_kubik = kubik_value * 273
            # print(f"Kubik: {kubik_value}")
            print(f"Kubik: {kubik_value}, Konversi ke kilo (x273): {round(konversi_kubik)}")
            total_berat_1 += konversi_kubik
    # else:
    #     print("Tidak ditemukan nilai setelah kata 'Kubik:' dalam teks.")

    if kilo_match:
        kilo_value = kilo_match.group(1)
        if kilo_value == '-':
            print(f"Kilo: {kilo_value}")
        else:
            kilo_value = int(float(kilo_value))
            print(f"Kilo: {kilo_value}")
            total_berat_1 += kilo_value
    # else:
    #     print("Tidak ditemukan nilai setelah kata 'Kilo:' dalam teks.")
    if total_berat_1 != 0:
        fill_total_berat_1 = total_berat_1
        print(f"Total Berat: {round(fill_total_berat_1)}")

    # =====BERAT=====END

    # =====TUJUAN====START

    tujuan = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[4]").text
    if alamat_pengirim is not None and alamat_penerima is not None:
        print(f"Tujuan: {tujuan}")
        fill_tujuan = tujuan

    # =====TUJUAN====END

list_detail_resi.append([fill_nama_pengirim, fill_nama_penerima, fill_nomor_pengirim, fill_nomor_penerima, fill_data_barang, fill_ongkir, round(fill_total_berat_1), fill_tujuan, fill_alamat_pengirim, fill_alamat_penerima])
print(list_detail_resi)

with open("C:\\Users\\Aditya PC\\Downloads\\Surat Jalan\\detail-resi.csv", "a", newline="") as csvfile:
    writer = csv.writer(csvfile)
    for b in list_detail_resi:
        writer.writerow(b)


#++++++++++++++LOOP PENCARIAN PERTAMA END+++++++++


# ///////////////////////////


#++++++LOOP PENCARIAN ITERASI START+++++++++

kolom_resi_iterasi = df['Resi'][1:].tolist()
baris_ke = 1
for i in kolom_resi_iterasi:
    # list_detail_resi = []
    re_input = driver.find_element(By.XPATH, '/html/body/div[3]/div/table[1]/tbody/tr/td[1]/form/table/tbody/tr/td[2]/input')

    button_2 = driver.find_element(By.XPATH, '/html/body/div[3]/div/table[1]/tbody/tr/td[1]/form/table/tbody/tr/td[3]/a/button')
    re_input.clear()  # Membersihkan nilai sebelum memasukkan yang baru
    re_input.send_keys(i)
    print(f"--------{i}--------")
    baris_ke += 1
    print(f"Resi ke: {baris_ke}/{len(kolom_resi_iterasi)}")
    button_2.click()  # Mengklik tombol

    history_barang_tabel = driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/table')
    history_barang_row = history_barang_tabel.find_elements(By.TAG_NAME, "tr")
    history_barang_row_count = len(history_barang_row)
    print(f"Jumlah baris history barang: {history_barang_row_count}")

    for e in range (1, history_barang_row_count):
        list_detail_resi = []
        sel_history_barang = driver.find_element(By.XPATH, f'/html/body/div[3]/div/div[2]/table/tbody/tr[{e}]/td[4]').text

        # print(sel_history_barang)


        # =====ALAMAT=====START
        # Regex Alamat
        pengirim_pattern = r"Alamat:\s(.*?)\s-"
        penerima_pattern = r"Alamat:\s(.*?)\s-.*?Penerima\s\(Alamat:\s(.*?)\s-"

        pengirim_match = re.search(pengirim_pattern, sel_history_barang)
        penerima_match = re.search(penerima_pattern, sel_history_barang)

        alamat_pengirim = None
        alamat_penerima = None
        if pengirim_match and penerima_match:
            alamat_pengirim = pengirim_match.group(1).replace(", ","_")
            alamat_penerima = penerima_match.group(2).replace(", ","_")
            if alamat_pengirim is not None and alamat_penerima is not None:
                fill_alamat_pengirim = alamat_pengirim
                fill_alamat_penerima = alamat_penerima
                
                print("Alamat Pengirim:", fill_alamat_pengirim)
                print("Alamat Penerima:", fill_alamat_penerima)

        # else:
        #     print("Tidak ada alamat pengirim dan penerima")

        # =====ALAMAT=====END


        # =====NAMA=====START

        nama_pengirim = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[2]").text
        nama_penerima = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[3]").text

        if alamat_pengirim is not None and alamat_penerima is not None:
            fill_nama_pengirim = nama_pengirim
            fill_nama_penerima = nama_penerima
            print(f"Nama Pengirim: {fill_nama_pengirim}")
            print(f"Nama Penerima: {fill_nama_penerima}")
        else:
            nama_pengirim = None
            nama_penerima = None

        # =====NAMA=====END


        # =====NO TLP=====START
        # Regex Nomor Telepon
        phone_numbers = re.findall(r'Telp: (\d+)', sel_history_barang)

        # Loop Nomor Telepon
        nomor_pengirim = None
        nomor_penerima = None
        if phone_numbers:
            nomor_pengirim = phone_numbers[0]
            nomor_penerima = phone_numbers[1]
            if nomor_pengirim is not None or nomor_penerima is not None:
                if nomor_pengirim == '-' or nomor_penerima == '-':
                    nomor_pengirim = '-'
                    nomor_penerima = '-'
                else:
                    fill_nomor_pengirim = nomor_pengirim
                    fill_nomor_penerima = nomor_penerima
                    print(f"Nomor Pengirim: {nomor_pengirim}")
                    print(f"Nomor Penerima: {nomor_penerima}")


        # =====NO TLP=====END


        # =====ISI=====START

        isi_barang_pattern = r"Isi:\s(.*?)\),"
        isi_barang_match = re.search(isi_barang_pattern, sel_history_barang)

        data_barang = None
        if isi_barang_match:
            data_barang = isi_barang_match.group(1)
            if data_barang is not None:
                if data_barang == '':
                    data_barang = "-"
                    fill_data_barang = data_barang
                    print(f"Isi Barang: {fill_data_barang}")
                else:
                    fill_data_barang = data_barang
                    print(f"Isi Barang:{fill_data_barang}")
        # else:
        #     print("Tidak dapat menemukan data barang.")

        # =====ISI=====END


        # =====ONGKIR=====START
        # Regex ongkir
        ongkir_pattern = r"Total:\sRp\.(\d{1,3}(?:,\d{3})*)"
        ongkir_match = re.search(ongkir_pattern, sel_history_barang)

        ongkir = None
        if ongkir_match:
            ongkir_string = ongkir_match.group(1)
            ongkir_string = ongkir_string.replace(",", "")
            ongkir = int(ongkir_string)
            if ongkir is not None:
                fill_ongkir = ongkir
                print("Ongkir:", fill_ongkir)
        # else:
        #     print("Tidak dapat menemukan ongkir.")

        # =====ONGKIR=====END

        # =====BERAT=====START

        kubik_pattern = r'Kubik: (-|\d+\.\d+)'
        kubik_match = re.search(kubik_pattern, str(sel_history_barang))

        kilo_pattern = r'Kilo: (\d+)'
        kilo_match = re.search(kilo_pattern, str(sel_history_barang))

        total_berat_1 = 0
        if kubik_match:
            kubik_value = kubik_match.group(1)
            if kubik_value == '-':
                print(f"Kubik: {kubik_value}")
            else:
                kubik_value = float(kubik_value)
                konversi_kubik = kubik_value * 273
                # print(f"Kubik: {kubik_value}")
                print(f"Kubik: {kubik_value}, Konversi ke kilo (x273): {round(konversi_kubik)}")
                total_berat_1 += konversi_kubik
        # else:
        #     print("Tidak ditemukan nilai setelah kata 'Kubik:' dalam teks.")

        if kilo_match:
            kilo_value = kilo_match.group(1)
            if kilo_value == '-':
                print(f"Kilo: {kilo_value}")
            else:
                kilo_value = int(float(kilo_value))
                print(f"Kilo: {kilo_value}")
                total_berat_1 += kilo_value
        # else:
        #     print("Tidak ditemukan nilai setelah kata 'Kilo:' dalam teks.")
        if total_berat_1 != 0:
            fill_total_berat_1 = total_berat_1
            print(f"Total Berat: {round(fill_total_berat_1)}")

        # =====BERAT=====END

        # =====TUJUAN====START

        tujuan = driver.find_element(By.XPATH, "/html/body/div[3]/div/table[2]/tbody/tr[2]/td[4]").text
        if alamat_pengirim is not None and alamat_penerima is not None:
            print(f"Tujuan: {tujuan}")
            fill_tujuan = tujuan

        # =====TUJUAN====END

    list_detail_resi.append([fill_nama_pengirim, fill_nama_penerima, fill_nomor_pengirim, fill_nomor_penerima, fill_data_barang, fill_ongkir, round(fill_total_berat_1), fill_tujuan, fill_alamat_pengirim, fill_alamat_penerima])
    print(list_detail_resi)

    with open("C:\\Users\\Aditya PC\\Downloads\\Surat Jalan\\detail-resi.csv", "a", newline="") as csvfile:
        writer = csv.writer(csvfile)
        for b in list_detail_resi:
            writer.writerow(b)

                
#++++++++++++++LOOP PENCARIAN ITERASI END+++++++++
