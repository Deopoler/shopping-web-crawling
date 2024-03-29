from PIL import Image
from urllib.request import urlopen
import yaml
import os
import numpy as np
import pandas as pd
import ctypes

with open("config.yaml", "r", encoding="utf8") as file:
    config = yaml.safe_load(file)


def center_crop(img, new_width=None, new_height=None):

    width = img.shape[1]
    height = img.shape[0]

    if new_width is None:
        new_width = min(width, height)

    if new_height is None:
        new_height = min(width, height)

    left = int(np.ceil((width - new_width) / 2))
    right = width - int(np.floor((width - new_width) / 2))

    top = int(np.ceil((height - new_height) / 2))
    bottom = height - int(np.floor((height - new_height) / 2))

    if len(img.shape) == 2:
        center_cropped_img = img[top:bottom, left:right]
    else:
        center_cropped_img = img[top:bottom, left:right, ...]

    return center_cropped_img


df = pd.read_excel(config["excel"], config["sheet"], dtype=str)

for i, row in df.iterrows():
    if not (type(row[config["image"]]) is str):
        continue
    url = row[config["image"]].replace(" ", "/")
    try:
        img = Image.open(urlopen(url))
    except:
        WS_EX_TOPMOST = 0x40000
        windowTitle = "에러"
        message = f"{row[config['filename']]}의 이미지를 불러올 수 없습니다."
        ctypes.windll.user32.MessageBoxExW(None, message, windowTitle, WS_EX_TOPMOST)
        continue

    img = img.rotate(config["rotate"], resample=Image.BICUBIC)
    size = img.size
    size_r = (size[0] * config["scale"] // 100, size[1] * config["scale"] // 100)
    img = img.resize(size_r, resample=Image.LANCZOS)
    img = Image.fromarray(center_crop(np.array(img), size[0], size[1]))
    if not os.path.exists("output/"):
        os.makedirs("output/")
    img.save(f"output/{row[config['filename']]}.jpg")
