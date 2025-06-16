import ast
import re
import pandas as pd
import ezdxf

def create_dxf_from_csv(csv_path):
    # CSV 읽기
    df = pd.read_csv(csv_path)


    # label 기준으로 그룹화
    grouped = df.groupby('label')

    for label, group_df in grouped:
        # index 순서대로 정렬 (만약 index 순서가 중요하다면)
        group_df = group_df.sort_values('index')

        # (CAD_x, CAD_y) 좌표 추출
        points = group_df[['CAD_x', 'CAD_y']].values  # shape = (N, 2)

        # 폴리곤 bounding box 계산
        min_x, max_x = points[:, 0].min(), points[:, 0].max()
        min_y, max_y = points[:, 1].min(), points[:, 1].max()

        width, height = max_x - min_x, max_y - min_y

        # 등각 스케일(Uniform Scale)로 540 x 540 박스에 맞추기
        if width == 0 or height == 0:
            scale = 1.0  # 혹은 다른 처리
        else:
            scale = min(540 / width, 540 / height)

        # 스케일링 후 좌표 변환
        transformed_points = [( (x - min_x) * scale, (y - min_y) * scale ) for x, y in points]

        # 스케일링 후 새로운 중심 좌표 계산
        min_x_t, max_x_t = min(p[0] for p in transformed_points), max(p[0] for p in transformed_points)
        min_y_t, max_y_t = min(p[1] for p in transformed_points), max(p[1] for p in transformed_points)

        center_x_t = (min_x_t + max_x_t) / 2
        center_y_t = (min_y_t + max_y_t) / 2

        # 중앙(270,270)으로 이동
        dx, dy = 270 - center_x_t, 270 - center_y_t
        final_points = [(x + dx, y + dy) for x, y in transformed_points]

        # DXF 문서 생성
        doc = ezdxf.new(dxfversion='R2010')
        msp = doc.modelspace()

        # LWPolyline 추가 (닫힌 다각형)
        msp.add_lwpolyline(final_points, close=True)

        # label 이름으로 파일 저장
        dxf_filename = f"{label}.dxf"
        doc.saveas(dxf_filename)
        print(f"Saved: {dxf_filename}")

# 사용 예시
if __name__ == "__main__":
    create_dxf_from_csv(r"C:\Users\user\Desktop\testf\테스트데이터.csv")
