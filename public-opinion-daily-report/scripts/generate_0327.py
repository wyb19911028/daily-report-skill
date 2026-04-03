#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成03月27日舆情监测日报 Word 文档
"""

import sys
sys.path.insert(0, '/workspace/projects/public-opinion-daily-report/scripts')

from create_docx import create_daily_overview, create_positive_info

# 日报概述内容
overview_content = {
    'title': '【03月27日机关重点舆情监测日报】',
    'time_range': '03月26日18:00-03月27日19:00',
    'overview': '传播概述：全网关于Y总、H总相关报道9167篇次，正面/中性8534篇占比93.09%；负面633篇占6.91%。正面新闻主要聚焦："华为春季全场景新品发布会首次在长沙举办，双方合作迈入全新阶段"、"重返千元机市场 华为冲刺鸿蒙生态"、"鸿蒙智行电池技术负责人：坚持使用电芯正置+壳体不带电方案"、"鸿蒙智行尚界Z7开启预售，22.98万元起，续航905km支持800V超充"、"鸿蒙智行V9亮相：2+2+3七座布局，增程版续航超200km！"等话题。负面新闻主要聚焦：""车还在厂里，配置已过时" 问界M7激光雷达换代，新车主都成了"大冤种"？"、"40万的问界M8开了俩月排气管全锈了 4S店出歪招：免费更换，怕贬值可不留维修记录"、"鸿蒙智行车辆智驾状态下与幼童发生碰撞，官方客服称"非车辆问题所致""等相关敏感内容。',
    'data': [
        '1、全网声量：全网关于Y总、H总相关报道为9167频次。',
        '2、正面声量：正面/中性信息8534频次，占93.09%。负面声量：负面信息633频次，占6.91%。',
        '3、评论监控：监测时间段内，用户评论共175698条，正面/中性占比93.23%（163796条），负面占比6.77%（11902条），未出现用户集中吐槽内容。'
    ]
}

# 高频正面信息内容
positive_content = {
    'positive_info': [
        {'title': '华为春季全场景新品发布会首次在长沙举办，双方合作迈入全新阶段', 'url': 'https://www.icswb.com/newspaper_article-detail-1836009.html'},
        {'title': '重返千元机市场 华为冲刺鸿蒙生态', 'url': 'https://epaper.nfnews.com/nfdaily/html/202603/27/content_10166326.html'},
        {'title': '鸿蒙智行电池技术负责人：坚持使用电芯正置+壳体不带电方案', 'url': 'https://news.mydrivers.com/1/1112/1112053.htm'},
        {'title': '鸿蒙智行尚界Z7开启预售，22.98万元起，续航905km支持800V超充', 'url': 'http://m.news18a.com/news/storys_242439.html'},
        {'title': '鸿蒙智行V9亮相：2+2+3七座布局，增程版续航超200km！', 'url': 'https://www.d1ev.com/newsflash/293202'}
    ]
}

# 生成文档
file1 = create_daily_overview('0327', overview_content)
file2 = create_positive_info('0327', positive_content)

print(f"已生成文档：")
print(f"1. {file1}")
print(f"2. {file2}")
