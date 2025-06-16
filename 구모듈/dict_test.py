import argparse
import os
import pandas as pd

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Read text from an image and filter by specific conditions")
    parser.add_argument('--number', type=str)
    args = parser.parse_args()

    number = args.number

    print('number', number)

    Dam_number = {
        "01": '01SYD-5YSR',
        "03": "03CJD-5YSR",
        "06": "06NGD-5YSR",
        "7": "07MYD-5YSR",
        "10": "10ADD-5YSR",
        "12": "12IHD-5YSR",
        "15": "15BRD-5YSR",
        "16": "16YDD-5YSR",
        "18": "18SJD-5YSR",
        "19": "19JHD-5YSR",
        "20": "20JAD-5YSR",
        "21": "21GDD-5YSR",
        "22": "22DBD-5YSR",
        "23": "23GPD-5YSR",
        "27": "27SYD-5YSR",
        "29": "29AGD-5YSR",
        "34": "34PRD-5YSR"
    }

    print(Dam_number[number])

    file = os.path.join('StateEstimatorReport', str(Dam_number[number])+'-231205.conf')

    try:
        with open(file, 'r', encoding='utf-8') as file:
            conf_content = file.read()
            print(conf_content)
    except FileNotFoundError:
        print(f"파일 '{file}'를 찾을 수 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


