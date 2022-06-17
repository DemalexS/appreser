from pprint import pprint
from bs4 import BeautifulSoup
import requests
import openpyxl
import statistics
 
wbname = 'Аналоги.xlsx'
def autoru_appraiser(wbname):
    wb = openpyxl.load_workbook(wbname)
    sheetanalog = wb['Аналоги']
    sheetobject = wb['Объекты оценки']
    columns = ['Марка',	'Модель','Год выпуска','Объем двигателя','Мощность двигателя','Тип двигателя',	'Привод',	'Тип кузова',	'Тип КПП',	'Цвет',	'Пробег',	'Цена',	'Ссылка', '', 'Цена с корректировкой на торг', 'Цена с корректировкой на год', 'Цена с коррестировкой на пробег', 'Валовая корректировка', 'Стоимость с учетом всех корректировок']
    iobj = 2

    #login = str(sheetobject.cell(row=1, column=14).value)
    #password = str(sheetobject.cell(row=1, column=16).value)
    #http_proxy  = "http://" + login + ":" + password + "@bcpsg-moscow.headoffice.psbank.local:8080"
    #https_proxy = "https://" + login + ":" + password + "@bcpsg-moscow.headoffice.psbank.local:8080"
    #proxyDict = {
    #              "http"  : http_proxy,
    #             "https" : https_proxy
    #          }

    while sheetobject.cell(row=iobj, column=1).value != '':
        objmarka = sheetobject.cell(row=iobj, column=1).value
        try:
            objmarka = objmarka.replace('-', '_')
            objmarka = objmarka.replace(' ', '_')
        except:
            pass
        objmodel = sheetobject.cell(row=iobj, column=2).value
        try:
            objmodel = objmodel.replace('-', '_')
            objmodel = objmodel.replace(' ', '_')
        except:
            pass
        objyear = sheetobject.cell(row=iobj, column=3).value
        objengvol = sheetobject.cell(row=iobj, column=4).value
        objhp = sheetobject.cell(row=iobj, column=5).value
        #objeng = sheetobject.cell(row=iobj, column=6).value
        #КПП
        if sheetobject.cell(row=iobj, column=9).value == 'робот':
            objkpp = 'ROBOT'
        elif sheetobject.cell(row=iobj, column=9).value == 'механика':
            objkpp = 'MECHANICAL'
        elif sheetobject.cell(row=iobj, column=9).value == 'автомат':
            objkpp = 'AUTOMATIC'
        elif sheetobject.cell(row=iobj, column=9).value == 'вариатор':
            objkpp = 'VARIATOR'
        #ПРИВОД
        if sheetobject.cell(row=iobj, column=7).value == 'полный':
            objgear = 'ALL_WHEEL_DRIVE'
        if sheetobject.cell(row=iobj, column=7).value == 'задний':
            objgear = 'REAR_DRIVE'
        if sheetobject.cell(row=iobj, column=7).value == 'передний':
            objgear = 'FORWARD_CONTROL'
        #Тип двигателя
        if sheetobject.cell(row=iobj, column=6).value == 'бензин':
            objeng = 'GASOLINE'
        elif sheetobject.cell(row=iobj, column=6).value == 'дизель':
            objeng = 'DIESEL'
        elif sheetobject.cell(row=iobj, column=6).value == 'гибрид':
            objeng = 'HYBRID'
        elif sheetobject.cell(row=iobj, column=6).value == 'электро':
            objeng = 'ELECTRO'
        #Тип кузова
        if sheetobject.cell(row=iobj, column=8).value == 'седан':
            objtob = 'SEDAN'
        elif sheetobject.cell(row=iobj, column=8).value == 'хэтчбек':
            objtob = 'HATCHBACK'
        elif sheetobject.cell(row=iobj, column=8).value == 'хэтчбек 3дв.':
            objtob = 'hatchback_3_doors'
        elif sheetobject.cell(row=iobj, column=8).value == 'хэтчбек 5дв.':
            objtob = 'hatchback_5_doors'
        elif sheetobject.cell(row=iobj, column=8).value == 'внедорожник':
            objtob = 'allroad'
        elif sheetobject.cell(row=iobj, column=8).value == 'внедорожник 3дв.':
            objtob = 'allroad_3_doors'
        elif sheetobject.cell(row=iobj, column=8).value == 'внедорожник 5дв.':
            objtob = 'allroad_5_doors'
        elif sheetobject.cell(row=iobj, column=8).value == 'универсал':
            objtob = 'wagon'
        elif sheetobject.cell(row=iobj, column=8).value == 'купе':
            objtob = 'coupe'
        elif sheetobject.cell(row=iobj, column=8).value == 'минивэн':
            objtob = 'minivan'
        elif sheetobject.cell(row=iobj, column=8).value == 'пикап':
            objtob = 'pickup'
        elif sheetobject.cell(row=iobj, column=8).value == 'лимузин':
            objtob = 'limousine'
        elif sheetobject.cell(row=iobj, column=8).value == 'фургон':
            objtob = 'van'
        elif sheetobject.cell(row=iobj, column=8).value == 'кабриолет':
            objtob = 'cabrio'
        #objgear = sheetobject.cell(row=iobj, column=7).value
        #objtob = sheetobject.cell(row=iobj, column=8).value
        #objkpp = sheetobject.cell(row=iobj, column=9).value
        objmileage = sheetobject.cell(row=iobj, column=11).value
        #print(objyear)
        try:
            year1 = objyear - 1
        except:
            break
        #print(year1, objyear)
        year2 = objyear + 1

        hp1 = round(float(objhp) * 0.95)
        hp2 = round(float(objhp) * 1.05)

        vol1 = round(int(objengvol) * 0.9,-2)
        if 3000 < vol1 < 3500:
            vol1 = 3000
        if 3500 < vol1 < 4000:
            vol1 = 3500
        if 4000 < vol1 < 4500:
            vol1 = 4000
        if 4500 < vol1 < 5000:
            vol1 = 4500
        if 5000 < vol1 < 5500:
            vol1 = 5000
        if 5500 < vol1 < 6000:
            vol1 = 5500
        if 6000 < vol1 < 7000:
            vol1 = 6000
        if 7000 < vol1 < 8000:
            vol1 = 7000
        if 8000 < vol1 < 9000:
            vol1 = 8000
        if 9000 < vol1:
            vol1 = 9000
        vol2 = round(int(objengvol) * 1.1,-2)
        if 3000 < vol2 < 3500:
            vol2 = 3500
        if 3500 < vol2 < 4000:
            vol2 = 4000
        if 4000 < vol2 < 4500:
            vol2 = 4500
        if 4500 < vol2 < 5000:
            vol2 = 5000
        if 5000 < vol2 < 5500:
            vol2 = 5500
        if 5500 < vol2 < 6000:
            vol2 = 6000
        if 6000 < vol2 < 7000:
            vol2 = 7000
        if 7000 < vol2 < 8000:
            vol2 = 8000
        if 8000 < vol2 < 9000:
            vol2 = 9000
        if 9000 < vol2 < 10000:
            vol2 = 10000

        millage1 = int(objmileage) * 0.75
        millage2 = int(objmileage) * 1.25
        if millage2 < 10000:
            millage2 = millage2 + 30000
        totalprice = 0
        totalpriceall = 0
        totalpricemedian = []

        url = 'https://auto.ru/cars/' + str(objmarka) + '/' + str(objmodel) + '/all/body-' + str(objtob) + '/?year_from=' + str(year1) + '&year_to=' + str(year2) + '&power_from=' + str(hp1) + '&displacement_from=' + str(vol1).replace('.0','') + '&displacement_to=' + str(vol2).replace('.0','') + '&transmission=' + str(objkpp) + '&power_to=' + str(hp2) + '&km_age_from=' + str(round(millage1,-1)).replace('.0','') + '&km_age_to=' + str(round(millage2,-1)).replace('.0','') + '&engine_group=' + str(objeng) + '&gear_type=' + str(objgear)
        #body_type_group = SEDAN & body_type_group = HATCHBACK & body_type_group = HATCHBACK_3_DOORS & body_type_group = HATCHBACK_5_DOORS & body_type_group = LIFTBACK & body_type_group = ALLROAD & body_type_group = ALLROAD_3_DOORS & body_type_group = ALLROAD_5_DOORS & body_type_group = WAGON & body_type_group = COUPE & body_type_group = MINIVAN & body_type_group = PICKUP & body_type_group = LIMOUSINE & body_type_group = VAN & body_type_group = CABRIO
        #ROBOT&transmission=AUTOMATIC&transmission=VARIATOR&transmission=AUTO
        #GASOLINE&engine_group=DIESEL&engine_group=HYBRID&engine_group=ELECTRO
        #FORWARD_CONTROL&gear_type=REAR_DRIVE&gear_type=ALL_WHEEL_DRIVE
        print(url)
        resp = requests.get(url)#, proxies = proxyDict, verify = False)
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        marka = soup.findAll('a', {'class' : 'Link BreadcrumbsGroup__itemText'})
        marka = marka[1].text
        model = soup.find('span', {'class' : 'BreadcrumbsGroup__itemText'}).text.replace(marka+' ','')
        car_list = soup.findAll('div', {'class' : 'ListingItem'})
        shname = 'Аналоги ' + str(iobj-1)
        try:
            delete = wb[shname]
            wb.remove(delete)
            wsanalog = wb.create_sheet(shname)
        except:
            wsanalog = wb.create_sheet(shname)
        #wsanalog = wb.create_sheet(shname)
        for iws, value in enumerate(columns, 1):
            wsanalog.cell(row=1, column=iws).value = value
        i = 2
        
        for car in car_list:
            try:
                carname = car.find('a', {'class' : 'Link ListingItemTitle__link'}).text
                caryear = car.find('div', {'class' : 'ListingItem__year'}).text
                cartech = car.findAll('div', {'class' : 'ListingItemTechSummaryDesktop__cell'})
                carlink = car.find('a', {'class' : 'Link OfferThumb'}).get('href')
                carmileage = car.find('div', {'class' : 'ListingItem__kmAge'}).text.replace(' км', '').replace(' ','')
                carkpp = cartech[1].text
                cartob = cartech[2].text
                cargear = cartech[3].text
                cartech_spl = cartech[0].text.split('/')
                carengvol = cartech_spl[0].replace(' л', '')
                carenghp = cartech_spl[1].replace(' л.с.', '')
                careng = cartech_spl[2].lower().replace(' ', '')
                carcolor = cartech[4].text
                carprice = car.findAll('span')
                try:
                    carprice = int(carprice[0].text.replace(' ₽','').replace(' ',''))
                except:
                    carprice = car.find('div', {'class' : 'ListingItemPrice__content'}).text
                    carprice = int(carprice.replace(' ₽', '').replace(' ', ''))
                wsanalog.cell(row=i, column=1).value = marka
                wsanalog.cell(row=i, column=2).value = model
                wsanalog.cell(row=i, column=3).value = int(caryear)
                wsanalog.cell(row=i, column=4).value = float(carengvol)
                wsanalog.cell(row=i, column=5).value = int(carenghp)
                wsanalog.cell(row=i, column=6).value = careng
                wsanalog.cell(row=i, column=7).value = cargear
                wsanalog.cell(row=i, column=8).value = cartob
                wsanalog.cell(row=i, column=9).value = carkpp
                wsanalog.cell(row=i, column=10).value = carcolor
                wsanalog.cell(row=i, column=11).value = int(carmileage)
                wsanalog.cell(row=i, column=12).value = int(carprice)
                wsanalog.cell(row=i, column=13).value = carlink
                wsanalog.cell(row=i, column=15).value = torgprice = int(int(carprice)*0.9)
                if int(objyear) == int(caryear):
                    wsanalog.cell(row=i, column=16).value = yearprice = int(torgprice)
                    yearcor = 0
                elif int(caryear) == int(objyear)-1:
                    wsanalog.cell(row=i, column=16).value = yearprice = int(int(torgprice)*1.07)
                    yearcor = -7
                elif int(caryear) == int(objyear)+1:
                    wsanalog.cell(row=i, column=16).value = yearprice = int(int(torgprice)*0.93)
                    yearcor = 7
                if int(objmileage)*0.95 < int(carmileage) <= int(objmileage)*0.98:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice) * 0.99)
                    millagecor = -1
                elif int(objmileage)*0.98 < int(carmileage) < int(objmileage)*1.02:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(yearprice)
                    millagecor = 0
                elif int(objmileage)*0.93 < int(carmileage) <= int(objmileage)*0.95:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.98)
                    millagecor = -2
                elif int(objmileage)*0.9 < int(carmileage) <= int(objmileage)*0.93:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.97)
                    millagecor = -3
                elif int(objmileage)*0.87 < int(carmileage) <= int(objmileage)*0.9:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.96)
                    millagecor = -4
                elif int(objmileage)*0.83 < int(carmileage) <= int(objmileage)*0.87:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.95)
                    millagecor = -5
                elif int(objmileage)*0.80 < int(carmileage) <= int(objmileage)*0.83:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.94)
                    millagecor = -6
                elif int(carmileage) <= int(objmileage)*0.80:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*0.98)
                    millagecor = -7
                elif int(objmileage)*1.02 <= int(carmileage) < int(objmileage)*1.05:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice) * 1.01)
                    millagecor = 1
                elif int(objmileage)*1.05 <= int(carmileage) < int(objmileage)*1.07:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.02)
                    millagecor = 2
                elif int(objmileage)*1.07 <= int(carmileage) < int(objmileage)*1.1:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.03)
                    millagecor = 3
                elif int(objmileage)*1.1 <= int(carmileage) < int(objmileage)*1.13:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.04)
                    millagecor = 4
                elif int(objmileage)*1.13 <= int(carmileage) < int(objmileage)*1.17:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.05)
                    millagecor = 5
                elif int(objmileage)*1.17 <= int(carmileage) < int(objmileage)*1.20:
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.06)
                    millagecor = 6
                elif int(objmileage)*1.20 <= int(carmileage):
                    wsanalog.cell(row=i, column=17).value = millageprice = int(int(yearprice)*1.07)
                    millagecor = 7
                wsanalog.cell(row=i, column=18).value = valcor = yearcor + millagecor
                wsanalog.cell(row=i, column=19).value = round(millageprice,-3)

                #print(i,year1, objyear)

                totalpriceall = totalpriceall + round(millageprice,-3)
                totalpricemedian.append(round(millageprice,-3))
                i = i + 1
                wb.save(wbname)
            except:
                pass
        wsanalog.cell(row=1, column=20).value = url
        wb.save(wbname)
        
        print('Отсев')
        ip = 2
        ic = i
        
        print(statistics.median(totalpricemedian))
        totalprice = round(totalpriceall / (ic - 2), -3)
        print(totalprice)
        while ip <= ic-1:
            #ПРИВОД
            #print(ip, 'привод', wsanalog.cell(row=ip, column=7).value.lower())
            if sheetobject.cell(row=iobj, column=7).value != wsanalog.cell(row=ip, column=7).value or sheetobject.cell(row=iobj, column=6).value != wsanalog.cell(row=ip, column=6).value.replace(' ', '') or hp1 > wsanalog.cell(row=ip, column=5).value or wsanalog.cell(row=ip, column=5).value > hp2:
                totalpriceall = totalpriceall - int(wsanalog.cell(row=ip, column=19).value)
                totalpricemedian.remove(int(wsanalog.cell(row=ip, column=19).value))
                wsanalog.delete_rows(ip, 1)
                print('удаляем', ip)
                wb.save(wbname)
                ic=ic-1
                ip = ip - 1
                        
            ip = ip + 1
            if wsanalog.cell(row=ip-1, column=19).value == '':
                break
        ip = 2
        try:
            totalprice = round(totalpriceall / (ic - 2), -3)
            print(statistics.median(totalpricemedian))
            print(totalprice)
            while ip <= ic-1:
                if int(wsanalog.cell(row=ip, column=19).value)/statistics.median(totalpricemedian) > 1.15 or int(wsanalog.cell(row=ip, column=19).value)/statistics.median(totalpricemedian) < 0.85:
                    totalpriceall = totalpriceall - int(wsanalog.cell(row=ip, column=19).value)
                    totalpricemedian.remove(int(wsanalog.cell(row=ip, column=19).value))
                    wsanalog.delete_rows(ip, 1)
                    wb.save(wbname)
                    ic=ic-1
                    totalprice = round(totalpriceall / (ic-2), -3)
                    print('удаляем', ip)
                    ip = ip - 1
                ip = ip + 1
                if wsanalog.cell(row=ip-1, column=19).value == '':
                    break
            ip = 2
            #ic = i
            
            totalprice = round(totalpriceall / (ic - 2), -3)
            print(statistics.median(totalpricemedian))
            print(totalprice)
            while ip <= ic - 1:
                if int(wsanalog.cell(row=ip, column=19).value) / statistics.median(totalpricemedian) > 1.17 or int(
                        wsanalog.cell(row=ip, column=19).value) / statistics.median(totalpricemedian) < 0.83:
                    totalpriceall = totalpriceall - int(wsanalog.cell(row=ip, column=19).value)
                    wsanalog.delete_rows(ip, 1)
                    wb.save(wbname)
                    ic = ic - 1
                    totalprice = round(totalpriceall / (ic - 2), -3)
                    print('удаляем', ip)
                    ip = ip - 1
                ip = ip + 1
                if wsanalog.cell(row=ip - 1, column=19).value == '':
                    break
            print(iobj-1, year1, totalprice, statistics.median(totalpricemedian))
            #totalprice = round(totalpriceall/(iс-2),-3)
        except:
            totalprice = 'Аналоги не найдены. Попробуйте изменить параметры ТС'
            
        sheetobject.cell(row=iobj, column=12).value = totalprice
        if sheetobject.cell(row=iobj, column=1).value == '':
            break
        iobj = iobj + 1
    wb.save(wbname)
#pprint.pp(cartech)

#<div class="ListingItem__kmAge">50&nbsp;км</div>
