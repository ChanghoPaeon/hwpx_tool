import logging
import pandas as pd

import configparser
config = configparser.ConfigParser() # 유니코드 문자열로 처리
with open('config.ini', 'r',encoding='utf-8') as f:
    config.read_file(f)

import datetime

suffix = datetime.datetime.now().strftime('%y%m%d_%H%M%S')

logging.basicConfig(filename="./log_file" + suffix + ".txt", level=logging.DEBUG,encoding='utf-8',
                    format="[ %(asctime)s | %(levelname)s ] %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S")

logger = logging.getLogger()

# logger.setLevel(logging.DEBUG)
# logger.setLevel(logging.INFO)
# file_handler=logging.FileHandler("./log_file.txt", encoding='utf-8')
# logger.addHandler(file_handler)


# def create_df():
#     df = pd.DataFrame({'col1': [1, 2, 3], 'col2': [4, 5, 6]})
#     df.to_csv('./df_test.csv', index=False)
#     return df
#
# logger.info("Dataframe 생성 완료")