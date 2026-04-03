#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成03月23日舆情监测日报 Word 文档
"""

import sys
sys.path.insert(0, '/workspace/projects/public-opinion-daily-report/scripts')

from create_docx import create_daily_overview, create_positive_info

# 日报概述内容
overview_content = {
    'title': '【03月23日机关重点舆情监测日报】',
    'time_range': '03月22日18:00-03月23日19:00',
    'overview': '传播概述：全网关于Y总、H总相关报道25255篇次，正面/中性24613篇占比97.46%；负面642篇占2.54%。正面新闻主要聚焦："华为何刚：原生鸿蒙设备数已突破5000万台"、"华为何刚：鸿蒙智行全系累计交付突破130万，连续14个月中国汽车品牌成交均价第一"、"麒麟芯片+鸿蒙6登上千元机，何刚：华为手机全面回归"、"华为余承东：尚界Z7/Z7T、问界M6/M7/M8、焕新享界双9、智界双7都将搭载896线双光路激光雷达"、"华为Mate80新版本首创超空间内存技术"等话题。负面新闻主要聚焦："何刚把余承东按在台下，顶替其发布鸿蒙智行新品，台下掌声雷动庆祝何总篡位成功"、"问界M7交付不足3个月就立马大升级 老车主深感"背刺"集体讨说法"、"智驾大赛新的比赛结果出炉，华为ADS依旧排名垫底"等相关敏感内容。',
    'data': [
        '1、全网声量：全网关于Y总、H总相关报道为25255频次。',
        '2、正面声量：正面/中性信息24613频次，占97.46%。负面声量：负面信息642频次，占2.54%。',
        '3、评论监控：监测时间段内，用户评论共170800条，正面/中性占比95.73%（163501条），负面占比4.27%（7299条），未出现用户集中吐槽内容。'
    ]
}

# 高频正面信息内容
positive_content = {
    'positive_info': [
        {'title': '华为何刚：原生鸿蒙设备数已突破5000万台', 'url': 'https://www.cnstock.com/commonDetail/654641'},
        {'title': '华为何刚：鸿蒙智行全系累计交付突破130万，连续14个月中国汽车品牌成交均价第一', 'url': 'https://www.ithome.com/0/931/782.htm'},
        {'title': '麒麟芯片+鸿蒙6登上千元机，何刚：华为手机全面回归', 'url': 'https://www.guancha.cn/economy/2026_03_23_811080.shtml'},
        {'title': '华为余承东：尚界Z7/Z7T、问界M6/M7/M8、焕新享界双9、智界双7都将搭载896线双光路激光雷达', 'url': 'https://www.ithome.com/0/931/619.htm'},
        {'title': '华为Mate80新版本首创超空间内存技术', 'url': 'https://www.stcn.com/article/detail/3691632.html'}
    ]
}

# 生成文档
file1 = create_daily_overview('0323', overview_content)
file2 = create_positive_info('0323', positive_content)

print(f"已生成文档：")
print(f"1. {file1}")
print(f"2. {file2}")
