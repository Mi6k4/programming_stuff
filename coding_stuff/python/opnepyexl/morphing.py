from openpyxl import Workbook, load_workbook

target_workbook = Workbook()
sheet=target_workbook.active

sheet["A1"]="Номер заказа у партнера"
sheet["B1"]="Номер сертификата" # номер сертификата
sheet["C1"]="Продукт в системе партнера" #назначение платежа
sheet["D1"]="Код продукта в BD"
sheet["E1"]="Дата начала действия" #дата оплаты
sheet["F1"]="Дата окончания действия" # дата окончания сертификата
sheet["G1"]="Стоимость"
sheet["H1"]="ФИО плательщика"
sheet["I1"]="Дата рождения плательщика"
sheet["J1"]="Пол плательщика"
sheet["K1"]="Номер телефона плательщика"
sheet["L1"]="Адрес электронной почты плательщика"
sheet["M1"]="Серия паспорта плательщика"
sheet["N1"]="Номер паспорта плательщика"
sheet["O1"]="Кем выдан паспорт плательщика"
sheet["P1"]="Дата выдачи паспорта плательщика"
sheet["Q1"]="Адрес плательщика"
sheet["R1"]="Гражданство плательщика"
sheet["S1"]="Город"
sheet["T1"]="Банк"
sheet["U1"]="Наименование ДО"

target_workbook.save(filename="target_table.xlsx")

morphing_workbook1=load_workbook(filename="sample_for_test.xlsx")
