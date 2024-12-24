import requests
import read_json_file
from read_json_file import load_config
import xlwings as xw
import random
import time
from database import store_data_to_db,get_comments_summary
from read_json_file import load_config,save_json_to_file
from read_json_file import response_path
import datetime as dt


config = load_config(read_json_file.json_path)
base_url = config['config']['base_url']
headers = config['config']['headers']


def fetch_to_review():
    print("开始执行爬取数据任务...")
    types = "new"
    all_review_data = []
    total_fetched = 0

    #通过 Excel 获取游戏 ID 和需要评论条数
    app_sheet = xw.Book.caller().sheets["数据录入"]
    app_id_value = app_sheet["app_id"].value
    feedback_cell = app_sheet["app_id"].offset(column_offset=1)
    number_value = app_sheet["number_value"].value
    feedback_cell2 = app_sheet["number_value"].offset(column_offset=1)

    # 检查并转换 app_id
    if not app_id_value:
        feedback_cell.value = "请提供你想要查询的appid"
        return
    try:
        app_id = int(app_id_value)
    except ValueError:
        feedback_cell.value = "appid 必须是一个有效的数字"
        return

    # 检查需要的评论数量
    if not number_value or number_value <= 0:
        feedback_cell2.value = "请提供一个有效的数量"
        return

    # 清除上次反馈结果
    try:
        if feedback_cell.api.MergeCells:
            feedback_cell.api.UnMerge()
        feedback_cell.clear_contents()

        if feedback_cell2.api.MergeCells:
            feedback_cell2.api.UnMerge()
        feedback_cell2.clear_contents()
    except Exception as e:
        print(f"清除内容时发生错误: {e}")

    try:
        while total_fetched < number_value:
            current_from_value = total_fetched + 1
            url = base_url.format(app_id, current_from_value, types)
            random_wait_time = random.uniform(3, 5)
            time.sleep(random_wait_time)
            print(f"请求URL: {url}")

            try:
                response = requests.get(url, headers=headers)
                print(f"响应状态码: {response.status_code}")

                if response.status_code == 200:
                    response.encoding = 'utf-8'
                    data = response.json()
                    review_list = data["data"]["list"]
                    current_fetched = len(review_list)
                    total_fetched += current_fetched
                    all_review_data.extend(review_list)
                    print(f"成功获取从 {current_from_value} 请求的评论，当前已获取 {total_fetched} 条评论")

                    if total_fetched >= number_value:
                        print(f"已获取{total_fetched}条评论，停止流程")
                        break
                else:
                    error_message = f"请求失败: {response.status_code} - {response.text}"
                    print(error_message)
                    feedback_cell.value = error_message
            except requests.exceptions.RequestException as e:
                error_message = f"请求发生异常: {str(e)}"
                print(error_message)
                feedback_cell.value = error_message

            time.sleep(5)
    except Exception as e:
        feedback_cell.value = f"Error: {str(e)}"
        return

    # 保存数据到文件并存入数据库,需要保存到本地取消注释
    # save_json_to_file({"data": {"list": all_review_data}}, response_path)
    store_data_to_db(all_review_data)
    check_database()

    # # # 显示任务完成信息
    feedback_cell.value = f"任务完成！共获取 {total_fetched} 条评论。"
    feedback_cell2.value = "数据已存入数据库并完成校验。"
    print("任务完成,请在Excel中查看结果。")


    
    
def check_database():
    sheet_db = xw.Book.caller().sheets["数据录入"]
    game_db = xw.Book.caller().sheets["制图所需数据（自动更新）"]
    game_db["game_list"].expand().clear_contents()
    sheet_db["log"].expand().clear_contents()

    
    # 获取当前 UTC 时间
    current_utc_time = dt.datetime.now(dt.timezone.utc)

    # 转换为东八区时间
    east_8_time = current_utc_time + dt.timedelta(hours=8)
    
    summary,game_names = get_comments_summary()
    sheet_db["log"].options(transpose=True).value = summary
    sheet_db["log"].expand().api.WrapText = False
    sheet_db["updated_at"].value = f"最后更新: {east_8_time.isoformat()}"
    game_db["game_list"].options(
        header=False, index=False).value= game_names
    