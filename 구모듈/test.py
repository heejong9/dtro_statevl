import ast
import glob
import os
import re
import pandas as pd
import ezdxf

def create_dxf_from_csv(csv_path,csvData):
   
    # CSV 읽기
    df = pd.read_csv(csv_path)
    csvData['Prefix'] =csvData['Defect_ID'].apply(lambda x: '_'.join(re.sub(r'^U\d+-', '', x).split('_')[:-1]))
    
    # 특정 Prefix로 Defect_ID 검색
    unique_prefixes = csvData['Prefix'].unique().tolist()
    
    # label 기준으로 그룹화
    grouped = df.groupby('label')
    
    for label, group_df in grouped:
        # index 순서대로 정렬 (만약 index 순서가 중요하다면)
        group_df = group_df.sort_values('index')
        
        # (CAD_x, CAD_y) 좌표 추출
        points = group_df[['CAD_x', 'CAD_y']].values  # shape = (N, 2)
        
        # 폴리곤 bounding box 계산
        min_x = points[:, 0].min()
        max_x = points[:, 0].max()
        min_y = points[:, 1].min()
        max_y = points[:, 1].max()
     
        width  = max_x - min_x
        height = max_y - min_y
        
        # 등각 스케일(Uniform Scale)로 540 x 540 박스에 맞추기
        # width 또는 height가 0에 가까운 경우 예외 처리 필요
        if width == 0 or height == 0:
            scale = 1.0  # 혹은 다른 처리
        else:
            scale = min(540 / width, 540 / height)
        
        # 스케일 + 평행이동(왼쪽 아래를 (0,0)에 맞춘 뒤 스케일링)
        transformed_points = []
        for (x, y) in points:
            x_t = (x - min_x) * scale
            y_t = (y - min_y) * scale
            transformed_points.append((x_t, y_t))
        
        # DXF 문서 생성
        doc = ezdxf.new(dxfversion='R2010')
        msp = doc.modelspace()
        
        # LWPolyline 추가 (닫힌 다각형)
        #  - close=True로 마지막 점과 첫 점을 연결해 닫힌 도형을 만듦
        msp.add_lwpolyline(transformed_points, close=True)
        
        
   
       
            
            
         
        
        
        # label 이름으로 파일 저장 (확장자는 .dxf)
        dxf_filename = f"{label}.dxf"
        doc.saveas(dxf_filename)
        print(f"Saved: {dxf_filename}")


# 사용 예시
if __name__ == "__main__":

    create_dxf_from_csv(r"C:\Users\user\Desktop\testf\테스트데이터.csv",csv)
