# taptap_review_inexcel
通过xlwings实现对taptap评论的拉取，存储、情感分析、词云&amp;可视化
----------------------------------------------------------------
配套python库：xlwings、pandas、sqlalchemy、snownlp、wordcloud
----------------------------------------------------------------
数据录入模块使用方法：
首先在数据录入sheet输入所要存入的游戏taptapid（可以通过taptap游戏界面的url获取）以及查询条数；
点击开始查询会开始爬取并存储到本地的review.db文件中
![image](https://github.com/user-attachments/assets/e5a78cbd-1f08-48f3-be6b-b74b0db3f5ed)

同时下方的查看数据库栏会显示当前数据库最后更新时间以及各游戏评论存储范围（会伴随开始查询自动更新，也可手动更新）
![image](https://github.com/user-attachments/assets/6d1b4e4c-04b2-4f7d-975f-07da2a9c55bb)

----------------------------------------------------------------

数据可视化主要分为三个模块：
评论条数（评分堆积柱状图）、评分情感度（小提琴图）、词云图
通过选择不同的指标，可以根据当前数据库数据生成对应图标可视化；
![image](https://github.com/user-attachments/assets/8288e277-8ef5-4e55-aa0f-5dab99b29f66)
![image](https://github.com/user-attachments/assets/fdecee69-878a-4df8-bcac-4d06ddd5b9ad)
![image](https://github.com/user-attachments/assets/1d91ee58-97c9-4654-8c60-42283c469766)

----------------------------------------------------------------
