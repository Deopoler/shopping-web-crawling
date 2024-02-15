import requests
import urllib
import yaml
import pandas as pd
import json
import ctypes

client_id = "your client id"
client_secret = "your client secret"

with open("config.yaml", "r", encoding="utf8") as file:
    config = yaml.safe_load(file)

category_df = pd.read_excel("category.xls", dtype=str)


def find_category(category1, category2, category3, category4):
    temp = category_df.iloc[
        category_df["대분류"]
        .searchsorted(category1, side="left") : category_df["대분류"]
        .searchsorted(category1, side="right")
    ]
    if category2 == "":
        return temp["카테고리코드"].tolist()[0]
    temp = temp.iloc[
        temp["중분류"]
        .searchsorted(category2, side="left") : temp["중분류"]
        .searchsorted(category2, side="right")
    ]
    if category3 == "":
        return temp["카테고리코드"].tolist()[0]
    temp = temp.iloc[
        temp["소분류"]
        .searchsorted(category3, side="left") : temp["소분류"]
        .searchsorted(category3, side="right")
    ]
    if category4 == "":
        return temp["카테고리코드"].tolist()[0]
    temp = temp.iloc[
        temp["세분류"]
        .searchsorted(category4, side="left") : temp["세분류"]
        .searchsorted(category4, side="right")
    ]
    return temp["카테고리코드"].tolist()[0]


df = pd.read_excel(config["excel"], config["sheet"], dtype=str)
result_df = pd.DataFrame(columns=["네이버카테고리", "네이버카테고리ID", "네이버상품명"])

for i, row in df.iterrows():
    if not (type(row[config["itemname"]]) is str):
        continue
    query = row[config["itemname"]]

    try:
        query = urllib.parse.quote(query)

        url = "https://openapi.naver.com/v1/search/shop?query=" + query

        request = urllib.request.Request(url)
        request.add_header("X-Naver-Client-Id", client_id)
        request.add_header("X-Naver-Client-Secret", client_secret)

        response = urllib.request.urlopen(request)
        data = json.load(response)
    except:
        WS_EX_TOPMOST = 0x40000
        windowTitle = "에러"
        message = f"{row[config['itemname']]}의 카테고리를 불러올 수 없습니다."
        ctypes.windll.user32.MessageBoxExW(None, message, windowTitle, WS_EX_TOPMOST)
        continue
    title = data["items"][0]["title"].replace("<b>", "").replace("</b>", "")
    category1 = data["items"][0]["category1"]
    category2 = data["items"][0]["category2"]
    category3 = data["items"][0]["category3"]
    category4 = data["items"][0]["category4"]
    categoryid = find_category(category1, category2, category3, category4)
    if category4 != "":
        result_df.loc[len(result_df)] = [
            f"{category1} > {category2} > {category3} > {category4}",
            categoryid,
            title,
        ]
    else:
        result_df.loc[len(result_df)] = [
            f"{category1} > {category2} > {category3}",
            categoryid,
            title,
        ]

result_df.to_excel("결과.xlsx")
