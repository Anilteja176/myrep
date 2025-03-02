import openpyxl
import openpyxl.worksheet
import requests
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active

sheet.append(["Name","Lower Price","Upper Price","Fuel","Mileage","Power"])
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
}

data = requests.get("https://www.cardekho.com/cars/Hyundai", headers=headers)

if data.status_code !=200:
    print(f"Failed to fetch the Webpage. Status Code: {data.status_code}")
else:
    soup = BeautifulSoup(data.text,"html.parser")
    cars = soup.find("body").find("div",class_="app-content").find("div",class_="gsc_row").\
        find_all("li",class_="gsc_col-xs-12 gsc_col-sm-6 gsc_col-md-12 gsc_col-lg-12")

    for car in cars:
        try:
            name = car.find("div",class_="gsc_col-sm-12 gsc_col-xs-12 gsc_col-md-8 listView holder posS").\
                find("h3").text
            L_price = car.find("div",class_="gsc_col-sm-12 gsc_col-xs-12 gsc_col-md-8 listView holder posS").\
                find("div",class_="price").text.split("*")[0].split("-")[0].split("Rs.")[1]
            U_price = car.find("div",class_="gsc_col-sm-12 gsc_col-xs-12 gsc_col-md-8 listView holder posS").\
                find("div",class_="price").text.split("*")[0].split("-")[1].split()[0]
            details1 = car.find("div",class_="gsc_col-sm-12 gsc_col-xs-12 gsc_col-md-8 listView holder posS").\
                find("div",class_="clearfix").find_all("div",class_="dotlist")[0].find_all("span")
            fuel = details1[0].text.strip()
            mileage = details1[1].text.strip() if fuel !="Electric" else details1[2].text.strip()
            power = details1[2]
            details2 = car.find("div",class_="gsc_col-sm-12 gsc_col-xs-12 gsc_col-md-8 listView holder posS").\
                find("div",class_="clearfix").find_all("div",class_="dotlist")[1].find_all("span")
            power = details2[1].text if fuel != "Electric" else details2[0].text.strip()

            print(name,L_price,U_price,fuel,mileage,power)
            sheet.append([name,L_price,U_price,fuel,mileage,power])
        except:
            continue
    save_path = r"E:\Excellenc\Projects\Demo1.xlsx"
    excel.save(save_path)
    print(f"Data saved to {save_path}")
