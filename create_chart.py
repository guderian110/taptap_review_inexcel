import xlwings as xw
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import matplotlib.dates as mdates
from database import get_chart_data
import xlwings as xw
import seaborn as sns
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import jieba
import jieba.posseg as pseg
from sentiments import analyze_sentiment
from collections import Counter
from pathlib import Path
import re


# 读取停用词列表
def load_stop_words(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return set(f.read().splitlines())
    
this_dir = Path(__file__).resolve().parent
stop_path = this_dir / "stop_words.txt"
dict_path = this_dir / "user_dict.txt"
   
# 去掉助词和语气词的词性列表
excluded_pos = {'u', 'y', 'p', 'r', 'e', 'o','c','d'}  # 'u': 助词, 'y': 语气词, 'p': 介词, 'r': 代词, 'e': 感叹词, 'o': 拟声词,'c': 连词,'d': 副词

def create_chart():
    wb = xw.Book.caller()
    chart_sheet = wb.sheets["数据可视化"]
    
    # 设置中文字体（如 SimHei 或者 Microsoft YaHei）
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文显示问题
    plt.rcParams['axes.unicode_minus'] = False    # 负号正常显示
    
    # 获取数据
    game_name = chart_sheet["game_name"].value
    feedback_cell_0 = chart_sheet["game_name"].offset(0, 1)
    argument = chart_sheet["argument"].value
    feedback_cell_1 = chart_sheet["argument"].offset(0, 1)
    days_number = chart_sheet["days_num"].value
    feedback_cell_2 = chart_sheet["days_num"].offset(0, 1)
    picture_cell = chart_sheet["latest_release"].offset(row_offset=2)
    datetime_value = chart_sheet["date_time"].value
    feedback_cell_3 = chart_sheet["date_time"].offset(0, 1)

    # 检查输入是否为空
    if not game_name:
        feedback_cell_0.value = "请选择你想要查询的游戏"
        return
    
    if not argument:
        feedback_cell_1.value = "请选择你想要查询的参数"
        return

    # 检查需要的评论数量
    if not days_number or days_number <= 0:
        feedback_cell_2.value = "请提供一个有效的数量"
        return
    
    if not datetime_value:
        feedback_cell_3.value = "请选择你想要倒退的日期"
        return

    # 删除已有的图像
    if "picture1" in chart_sheet.pictures:
        chart_sheet.pictures["picture1"].delete()
    if "frequency_chart" in chart_sheet.pictures:
        chart_sheet.pictures["frequency_chart"].delete()

    if argument == "评分":
        # 获取数据并处理
        data = get_chart_data(game_name, argument, days_number, datetime_value)
        data_long = data.melt(id_vars='日期', var_name='评分',value_name='评论数')
        
        # 强制评分列为有序类别，确保堆叠顺序
        score_order = [1, 2, 3, 4, 5]  # 底层为1，最高层为5
        data_long['评分'] = pd.Categorical(data_long['评分'], categories=score_order, ordered=True)

        # 排序数据，保证评分从低到高
        data_long = data_long.sort_values(['日期', '评分'])
        
        # 自定义高对比度颜色
        custom_palette = sns.color_palette("bright", len(score_order))
            
        # 绘制堆积柱状图
        plt.figure(figsize=(10, 6))
        
        plot=sns.histplot(
        data=data_long,
        x='日期',
        weights='评论数',
        hue='评分',
        hue_order=[1, 2, 3, 4, 5], 
        multiple="stack",  # 实现堆积
        element="bars",
        shrink=0.8 , # 控制柱子的宽度
        palette=custom_palette
        )
        
        # 显示所有横坐标标签
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))  # 格式化日期
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))  # 每天显示

        # 旋转日期标签
        plt.xticks(rotation=45, fontsize=10)

        # 设置图例位置，防止裁剪
        plt.title(f"{game_name} 每日评论堆积数量图", fontsize=16) 
        plt.xlabel("日期", fontsize=12)
        plt.ylabel("评论数量", fontsize=12)
        
        # 修复图例，移动到适当位置
        sns.move_legend(plot, "upper left", bbox_to_anchor=(1.05, 1), title="评分等级")

        # 自动调整布局
        plt.tight_layout()

        # 保存图形为文件
        plt.savefig('chart.png', dpi=300, bbox_inches='tight')
        plt.close()  # 关闭图形以释放资源
            # 将图形插入到 Excel 中
        chart_sheet.pictures.add('chart.png', name="picture1",
                                top=picture_cell.top,
                                left=picture_cell.left)
    elif argument == "评论情感度":
        
            data = get_chart_data(game_name, argument, days_number, datetime_value)
            data['sentiment_score'] = data['评论内容'].apply(lambda x: analyze_sentiment(x))
            
            data['sentiment_score'] = data['sentiment_score'].apply(
            lambda x: sum(x) / len(x) if isinstance(x, list) else x
                )

            # 将日期转换为字符串类型
            data['日期'] = data['日期'].astype(str)
         
            # 绘制小提琴图
            plt.figure(figsize=(12, 6))
            sns.violinplot(x='日期', y='sentiment_score', data=data, inner='box')
            plt.ylabel('情感得分 (0 到 1), 越接近 1 越正面')
            plt.title('各日期的评论情感得分分布')
            plt.xticks(rotation=45)
            plt.ylim(-1, 1.5)

            plt.tight_layout()
            plt.savefig('chart.png', dpi=300, bbox_inches='tight')
            plt.close()  # 关闭图形以释放资源

            # 将图形插入到 Excel 中
            chart_sheet.pictures.add('chart.png', name="picture1",
                                    top=picture_cell.top,
                                    left=picture_cell.left)
    elif argument == "词云图":
        data = get_chart_data(game_name, argument, days_number, datetime_value)
        jieba.load_userdict(str(dict_path))
         # 合并评论内容为一个字符串
        comments = ' '.join(data['评论内容'].astype(str))
        # 去掉中文符号
        comments = re.sub(r'[，。！？、；：“”‘’【】（）()《》]', '', comments)
        # 使用jieba进行分词和词性标注
        words_with_pos = pseg.cut(comments)
        # 使用jieba进行分词和词性标注
        stop_words = load_stop_words(str(stop_path))
        # 过滤掉停用词
        filtered_words = [word for word, flag in words_with_pos if word not in stop_words and len(word) >= 2 and flag not in excluded_pos]
        filtered_words = ' '.join(filtered_words)
        
        
        # 生成词云
        wordcloud = WordCloud(font_path=r'C:\Windows\Fonts\simhei.ttf', width=800, height=400, background_color='white',max_words=50).generate(filtered_words )
        
        # 绘制词云图
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title(f"{game_name} 词云图", fontsize=16)
        plt.tight_layout()
        plt.savefig('wordcloud.png', dpi=300, bbox_inches='tight')
        plt.close()

        # 统计高频词
        word_list = filtered_words .split()
        word_counts = Counter(word_list)
        most_common_words = word_counts.most_common(10)

         # 将高频词数据转换为 DataFrame 以便绘图
        common_words_df = pd.DataFrame(most_common_words, columns=['词', '频次'])

        # 绘制高频词条形图
        plt.figure(figsize=(10, 5))
        sns.barplot(x='频次', y='词', data=common_words_df, palette='viridis')
        plt.title(f"{game_name} 高频词出现频次图", fontsize=16)
        plt.xlabel("频次", fontsize=12)
        plt.ylabel("词", fontsize=12)
        plt.tight_layout()
        plt.savefig('frequency_chart.png', dpi=300, bbox_inches='tight')
        plt.close()

        # 将词云和频次图插入到同一个区域
        chart_sheet.pictures.add('wordcloud.png', name="picture1",
                                  top=picture_cell.top,
                                  left=picture_cell.left)

        # 计算频次图的插入位置，保持上下排列
        frequency_picture_cell = picture_cell.offset(row_offset=picture_cell.height + 10)
        chart_sheet.pictures.add('frequency_chart.png', name="frequency_chart",
                                  top=frequency_picture_cell.top,
                                  left=frequency_picture_cell.left)
    else:
        print("请输入正确的图表类型")