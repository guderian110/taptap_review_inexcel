from pathlib import Path
import sqlalchemy
from sqlalchemy.engine import Engine
from sqlalchemy import create_engine,Column, String, Integer, DateTime,func
from sqlalchemy import event
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import re
import datetime
import xlwings as xw
import pandas as pd



@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

this_dir = Path(__file__).resolve().parent
db_path = this_dir / "review.db"
   
Base = declarative_base()

engine = sqlalchemy.create_engine(f"sqlite:///{db_path}")
Base.metadata.create_all(engine)


class TapTapData(Base):
    __tablename__ = 'taptap_data'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    app_id = Column(String)
    game_name = Column(String)
    review_content = Column(String)
    review_time = Column(DateTime)
    rank = Column(Integer)
    score = Column(Integer)
    

def clean_text(text):
    # 去掉 'br' 字符
    text = re.sub(r'br', '', text)
    # 只保留中文字符、英文、数字和常见标点符号
    text = re.sub(r'[^\w\s,.!?？。！；：，、‘’“”【】（）()《》\u4e00-\u9fa5]', '', text)
    # 替换多余空格
    text = re.sub(r'\s+', ' ', text).strip()
    return text
        
def store_data_to_db(data):
    """
    批量存储数据到数据库，并检查重复项。
    """
    Session = sessionmaker(bind=engine)
    session = Session()
    
    review_data = []
    sheet_db = xw.Book.caller().sheets["数据录入"]  # 引用工作表用于写入错误信息

    try:
        for index, game in enumerate(data, start=1):
            app = game["moment"]["app"]
            app_name = app["title"]
            app_id = app["id"]
            review_content = clean_text(game["moment"]["review"]["contents"]["text"])
            score = game["moment"]["review"]["score"]
            review_time = datetime.datetime.fromtimestamp(game["moment"]["created_time"])

            # 检查数据库中是否已存在相同的记录
            existing_entry = session.query(TapTapData).filter(
                TapTapData.app_id == app_id,
                TapTapData.review_content == review_content,
                TapTapData.review_time == review_time
            ).first()

            if existing_entry:
                print(f"记录已存在，跳过插入: {existing_entry.game_name} - {existing_entry.review_time}")
                continue

            # 如果不存在，则创建新的实例
            review_entry = TapTapData(
                app_id=app_id,
                game_name=app_name,
                review_content=review_content,
                review_time=review_time,
                rank=index,
                score=score  
            )
            review_data.append(review_entry)
        
        # 批量添加并提交
        if review_data:  # 只有在有新数据时才提交
            session.add_all(review_data)
            session.commit()
            print(f"成功插入{len(review_data)}条新记录")
        else:
            print("没有新数据需要插入")
    except Exception as e:
        session.rollback()  # 回滚事务
        error_message = f"发生错误: {str(e)}"  # 获取错误信息
        print(error_message)
        sheet_db["log"].value = error_message  # 将错误信息写入 Excel 的 log 单元格
    finally:
        session.close() 

#当前数据库已有游戏数据拉取打印为日志    
def get_comments_summary():
    Session = sessionmaker(bind=engine)
    session = Session()
    
    results = session.query(
        TapTapData.game_name,
        func.count(TapTapData.review_content).label('comment_count'),
        func.min(TapTapData.review_time).label('min_time'),
        func.max(TapTapData.review_time).label('max_time')
    ).group_by(TapTapData.game_name).all()
  
    session.close()
    logs = []
    game_names = []
    
    for game_name, comment_count, min_time, max_time in results:
        logs.append(f"游戏名称: {game_name}")
        logs.append(f"评论数量: {comment_count}")
        logs.append(f"时间范围: {min_time} 到 {max_time}\n")
        game_names.append(game_name)
    return logs,game_names



#获取制图所需数据data
def get_chart_data(game_name,argument,number,datetime):
    Session = sessionmaker(bind=engine)
    session = Session()
    number = int(number)
    
    if isinstance(datetime, str):
        datetime = pd.to_datetime(datetime)  
    datetime = datetime.date()  
    
    # 根据参数选择适当的聚合函数  
    if argument == "评论条数":
    # 按日期和评分等级分组并计算每个等级的人数
        query = session.query(
            TapTapData.review_time,
            TapTapData.score
        ).filter(TapTapData.game_name == game_name) \
         .order_by(TapTapData.review_time.desc()) 
         
        result = query.all()
    
        df = pd.DataFrame(result, columns=['日期', '评分等级'])
        df['日期'] = pd.to_datetime(df['日期']).dt.date

        # 计算开始日期
        start_date = datetime - pd.Timedelta(days=number)
        df = df[(df['日期'] >= start_date) & (df['日期'] <= datetime)]
        print("筛选后的数据：")
        print(df)
    
        pivot_df = df.pivot_table(index='日期', columns='评分等级', aggfunc='size', fill_value=0)
        
        print("透视表数据：")
        print(pivot_df)
        
        # 关闭会话
        session.close()
        return pivot_df.reset_index()
    elif argument == "评论情感度":
        query = session.query(
            TapTapData.review_time,
            TapTapData.review_content
        ).filter(TapTapData.game_name == game_name) \
            .order_by(TapTapData.review_time.desc())
            
        result = query.all()

        df = pd.DataFrame(result, columns=['日期', '评论内容'])
        df['日期'] = pd.to_datetime(df['日期']).dt.date
        
        # 计算开始日期
        start_date = datetime - pd.Timedelta(days=number)
        df = df[(df['日期'] >= start_date) & (df['日期'] <= datetime)]
        print("筛选后的数据：")
        print(df)
        
        grouped_df = df.apply(list).reset_index()
        
        # 关闭会话
        session.close()
        return grouped_df
    elif argument == "词云图":
        query = session.query(
            TapTapData.review_time,
            TapTapData.review_content
        ).filter(TapTapData.game_name == game_name) \
            .order_by(TapTapData.review_time.desc())

        result = query.all()

        df = pd.DataFrame(result, columns=['日期', '评论内容'])
        df['日期'] = pd.to_datetime(df['日期']).dt.date

        # 计算开始日期
        start_date = datetime - pd.Timedelta(days=number)
        df = df[(df['日期'] >= start_date) & (df['日期'] <= datetime)]
        
        print("筛选后的数据：")
        print(df)
    
        # 关闭会话
        session.close()
        return df
    else:
        session.close()
        return "无效的参数"