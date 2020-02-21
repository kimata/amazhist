#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Usage: amazgraph.py [-h] [-j <json_path>]

Amazon の購入履歴データから，時系列グラフを生成します．

Options:
  -j <json_path>    履歴情報存を記録した JSON ファイルのパス  (required) [default: amazhist.json]
"""
from docopt import docopt

import json

import matplotlib

matplotlib.use("Agg")

from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt

import numpy as np
import re
from functools import partial

import os
import sys
import pprint

FONT_REGULAR_FILE = "./font/mplus-2p-regular.ttf"
FONT_BOLD_FILE = "./font/mplus-2p-bold.ttf"

PRICE_HIST_MAX = 20000
PRICE_HIST_BIN = 20
SUBCATEGORY_NUM = 5


def get_item_year(item):
    return int(re.search(r"^\d{4}", item["date"]).group())


def get_item_category(item):
    category = item["category"]

    if category == "":
        category = "?"

    return category


def get_item_subcategory(item):
    subcategory = item["subcategory"]

    if subcategory == "":
        subcategory = "?"

    return subcategory


def get_yeaer_range(item_list):
    year_set = set()
    for item in item_list:
        year_set.add(get_item_year(item))

    return list(range(sorted(year_set)[0], sorted(year_set)[-1] + 1, 1))


def get_by_year_price_hist(item_list, year_list):
    price_list_map = {year: [] for year in year_list}
    price_hist_map = {year: [] for year in year_list}

    for item in item_list:

        year = get_item_year(item)
        price = item["price"]

        if price > PRICE_HIST_MAX:
            price = PRICE_HIST_MAX

        price_list_map[year].append(price)

    for year in price_list_map.keys():
        price_hist_map[year] = np.histogram(
            price_list_map[year], bins=PRICE_HIST_BIN, range=(0, PRICE_HIST_MAX)
        )
    return price_hist_map


def get_category_list(item_list):
    category_map = {}
    for item in item_list:
        category = get_item_category(item)

        if category not in category_map:
            category_map[category] = 1
        else:
            category_map[category] += 1

    # NOTE: 数量が全体に締める割合が 1% 未満のカテゴリは捨てる
    category_list = filter(
        lambda category: (float(category_map[category]) / len(item_list)) > 0.01,
        category_map.keys(),
    )

    # NOTE: プロットしたときに見やすくなるように，総量が多い順でソート
    return sorted(
        category_list, key=lambda category: category_map[category], reverse=True
    )


def get_by_year_category_count(item_list, year_list, category_list):
    category_count_map = {
        year: {category: 0 for category in category_list} for year in year_list
    }

    for item in item_list:
        year = int(re.search(r"^\d{4}", item["date"]).group())
        category = get_item_category(item)

        if category in category_count_map[year]:
            category_count_map[year][category] += 1

    return category_count_map


def get_subcategory_list(item_list, category):
    subcategory_map = {}
    for item in item_list:
        if get_item_category(item) != category:
            continue

        subcategory = get_item_subcategory(item)

        if subcategory not in subcategory_map:
            subcategory_map[subcategory] = 1
        else:
            subcategory_map[subcategory] += 1

    # NOTE: 数量が全体に締める割合が 1% 未満のカテゴリは捨てる
    subcategory_list = filter(
        lambda subcategory: (float(subcategory_map[subcategory]) / len(item_list))
        > 0.01,
        subcategory_map.keys(),
    )

    # NOTE: プロットしたときに見やすくなるように，総量が多い順でソート
    return sorted(
        subcategory_list,
        key=lambda subcategory: subcategory_map[subcategory],
        reverse=True,
    )


def get_by_year_subcategory_count(item_list, category, year_list, subcategory_list):
    subcategory_count_map = {
        year: {subcategory: 0 for subcategory in subcategory_list} for year in year_list
    }

    for item in item_list:
        if get_item_category(item) != category:
            continue

        year = int(re.search(r"^\d{4}", item["date"]).group())
        subcategory = get_item_subcategory(item)

        if subcategory in subcategory_count_map[year]:
            subcategory_count_map[year][subcategory] += 1

    return subcategory_count_map


def create_graph_impl(
    file_path,
    title,
    xpos,
    ypos,
    zpos,
    dx,
    dy,
    dz,
    xticks,
    xlabel,
    ylabel,
    zlabel,
    func=None,
):
    fig = plt.figure(figsize=(12.0, 10.0), dpi=300)
    ax = fig.add_subplot(111, projection="3d")

    font_regular = matplotlib.font_manager.FontProperties(fname=FONT_REGULAR_FILE)
    font_bold = matplotlib.font_manager.FontProperties(
        fname=FONT_BOLD_FILE, weight="bold"
    )

    ax.set_title(title, fontproperties=font_bold, fontsize=18)

    if xlabel is not None:
        ax.set_xlabel(xlabel, labelpad=25, fontproperties=font_bold, fontsize=12)

    ax.set_ylabel(ylabel, labelpad=15, fontproperties=font_bold, fontsize=12)
    ax.set_zlabel(zlabel, labelpad=15, fontproperties=font_bold, fontsize=12)

    # NOTE: Y 軸の目盛は整数
    ax.get_yaxis().set_major_locator(matplotlib.ticker.MaxNLocator(integer=True))

    # NOTE: X 軸の値によって色を変える
    x_off = xpos - np.abs(min(xpos))
    colors = matplotlib.cm.hsv(x_off.astype(float) / x_off.max())

    # NOTE: 値が 0 の場合は表示しない
    mask = dz > 0.0

    ax.bar3d(
        xpos[mask],
        ypos[mask],
        zpos[mask],
        dx,
        dy,
        dz[mask],
        color=colors[mask],
        zsort="max",
    )

    ax.view_init(elev=60, azim=45)

    if func is not None:
        func(plt, ax, font_regular)

    plt.savefig(file_path)


def create_by_year_price_hist_graph(file_path, item_list, year_list):
    by_year_price_hist = get_by_year_price_hist(item_list, year_list)

    xpos, ypos = np.meshgrid(next(iter(by_year_price_hist.values()))[1][1:], year_list)

    xpos = xpos.flatten()
    ypos = ypos.flatten() - 0.5
    zpos = np.zeros_like(xpos)

    dx = (PRICE_HIST_MAX / PRICE_HIST_BIN) * 0.5
    dy = 0.5
    dz = np.array(
        list(map(lambda hist: hist[0], by_year_price_hist.values()))
    ).flatten()

    def custom_format(plt, ax, font):
        ax.xaxis.set_major_formatter(
            plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x)))
        )
        plt.xticks(
            range(
                int(PRICE_HIST_MAX / PRICE_HIST_BIN * 2),
                PRICE_HIST_MAX + int(PRICE_HIST_MAX / PRICE_HIST_BIN * 2),
                int(PRICE_HIST_MAX / PRICE_HIST_BIN * 2),
            ),
            ha="left",
            fontproperties=font,
        )

    create_graph_impl(
        file_path,
        "Amazon での購入価格ヒストグラムの時系列変化",
        xpos,
        ypos,
        zpos,
        dx,
        dy,
        dz,
        None,
        "価格",
        "年",
        "数量",
        custom_format,
    )


def create_by_year_category_count_graph(file_path, item_list, year_list, category_list):
    by_year_category_count = get_by_year_category_count(
        item_list, year_list, category_list
    )

    xpos, ypos = np.meshgrid(range(0, len(category_list)), year_list)

    xpos = xpos.flatten()
    ypos = ypos.flatten() - 0.5
    zpos = np.zeros_like(xpos)

    dx = 0.5
    dy = 0.5
    dz = np.array(
        [
            [by_year_category_count[year][category] for category in category_list]
            for year in year_list
        ]
    ).flatten()

    def custom_format(plt, ax, font):
        plt.xticks(
            range(0, len(category_list)),
            category_list,
            ha="left",
            rotation=-40,
            fontproperties=font,
        )

    create_graph_impl(
        file_path,
        "Amazon での購入カテゴリの時系列変化",
        xpos,
        ypos,
        zpos,
        dx,
        dy,
        dz,
        category_list,
        None,
        "年",
        "数量",
        custom_format,
    )


def create_by_year_subcategory_count_graph(file_path, item_list, year_list, category):
    subcategory_list = get_subcategory_list(item_list, category)
    by_year_subcategory_count = get_by_year_subcategory_count(
        item_list, category, year_list, subcategory_list
    )

    xpos, ypos = np.meshgrid(range(0, len(subcategory_list)), year_list)

    xpos = xpos.flatten()
    ypos = ypos.flatten() - 0.5
    zpos = np.zeros_like(xpos)

    dx = 0.5
    dy = 0.5
    dz = np.array(
        [
            [
                by_year_subcategory_count[year][subcategory]
                for subcategory in subcategory_list
            ]
            for year in year_list
        ]
    ).flatten()

    def format_xaxis(plt, ax, font_prop):
        plt.xticks(
            range(0, len(subcategory_list)),
            subcategory_list,
            ha="left",
            rotation=-40,
            fontproperties=font_prop,
        )

    create_graph_impl(
        file_path,
        "Amazon での購入カテゴリ({0:s})の時系列変化".format(category),
        xpos,
        ypos,
        zpos,
        dx,
        dy,
        dz,
        subcategory_list,
        None,
        "年",
        "数量",
        format_xaxis,
    )


def error(message):
    print("\033[31m{0:s}\033[0m".format(message))
    sys.exit(-1)


def create_graph(json_path):
    print("購入履歴デーグを読み込んでいます ...")

    if not os.path.exists(json_path):
        error("「{0:s}」は存在しません．amazhist.rb を実行して生成してください．".format(json_path))

    item_list = json.load(open(json_path, "r"))
    print("    - 履歴件数: {0:,}".format(len(item_list)))

    # NOTE: 購入が無い年もグラフ化したいので，購入履歴がある年の上下限を取得する
    year_list = get_yeaer_range(item_list)
    print("    - 年範囲  : {0}-{1}".format(year_list[0], year_list[-1]))

    category_list = get_category_list(item_list)
    print("    - カテゴリ: {0}".format(len(category_list)))

    print("")
    print("グラフを生成しています ", end="")
    create_by_year_price_hist_graph("時系列変化_購入価格.png", item_list, year_list)
    print(".", end="", flush=True)
    create_by_year_category_count_graph(
        "時系列変化_購入カテゴリ.png", item_list, year_list, category_list
    )
    print(".", end="", flush=True)
    for i in category_list[0:SUBCATEGORY_NUM]:
        create_by_year_subcategory_count_graph(
            "時系列変化_購入サブカテゴリ({0:s})_数量.png".format(i), item_list, year_list, i
        )
        print(".", end="", flush=True)
    print("")
    print("完了しました.")


create_graph(docopt(__doc__).get("-j"))

# Local Variables:
# coding: utf-8
# mode: python
# tab-width: 4
# indent-tabs-mode: nil
# End:
