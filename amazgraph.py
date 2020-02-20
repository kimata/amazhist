#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json

import matplotlib
matplotlib.use('Agg')
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt

import numpy as np
import re

PRICE_HIST_MAX = 20000
PRICE_HIST_BIN = 20
 
def get_yeaer_range(tem_list):
    year_set = set()
    for item in item_list:
        year_set.add(int(re.search(r'^\d{4}', item['date']).group()))

    return list(range(sorted(year_set)[0], sorted(year_set)[-1]+1, 1))

def get_by_year_price_hist(item_list, year_list):
    price_list_map = { year: [] for year in year_list }
    price_hist_map = { year: [] for year in year_list }

    for item in item_list:
        year = int(re.search(r'^\d{4}', item['date']).group())
        price = item['price']

        if (price > PRICE_HIST_MAX):
            price = PRICE_HIST_MAX

        price_list_map[year].append(price)

    for year in price_list_map.keys():
        price_hist_map[year] = np.histogram(price_list_map[year],
                                            bins=PRICE_HIST_BIN,
                                            range=(0,PRICE_HIST_MAX))
    return price_hist_map

def create_graph(file_path, item_list, year_list):
    by_year_price_hist = get_by_year_price_hist(item_list, year_list)

    fig = plt.figure(figsize=(12.0, 10.0))
    ax = fig.add_subplot(111, projection='3d')

    xpos, ypos = np.meshgrid(next(iter(by_year_price_hist.values()))[1][1:],
                             year_list)

    xpos = xpos.flatten()
    ypos = ypos.flatten() - 0.5
    zpos = np.zeros_like(xpos)

    dx = (PRICE_HIST_MAX/PRICE_HIST_BIN) *0.5* np.ones_like(zpos)
    dy = np.ones_like(zpos) * 0.5
    dz = np.array(list(map(lambda hist: hist[0], by_year_price_hist.values()))).flatten()

    y_off = ypos - np.abs(min(ypos))
    colors = matplotlib.cm.hsv(y_off.astype(float)/y_off.max())

    ax.set_title('Time series of price histgram', fontsize=18)

    ax.set_ylabel('Year', fontsize=12)
    ax.set_xlabel('Price', fontsize=12)
    ax.set_zlabel('Count', fontsize=12)

    ax.get_yaxis().set_major_locator(matplotlib.ticker.MaxNLocator(integer=True))

    ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=colors, zsort='max', shade=True)
    
    plt.savefig(file_path)



    
item_list = json.load(open('amazhist-200217.json', 'r'))
year_list = get_yeaer_range(item_list)
create_graph('a.png', item_list, year_list)




# Local Variables:
# coding: utf-8
# mode: python
# tab-width: 4
# indent-tabs-mode: nil
# End:
