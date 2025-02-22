#데이터 전처리
import pandas as pd
import numpy as np
import os

def read_excel_files(directory_path):
    """
    지정된 디렉토리 내의 모든 엑셀 파일을 읽고 기본 전처리를 수행
    - 첫 번째 row를 header로 사용
    - 중복된 X [s] 열 제거
    - GYRO, EMG. IMU 데이터 구분 

    Returns:
        dict: {파일명: 처리된 데이터프레임} 형태의 딕셔너리
    """
    excel_files = {}
    
    # 디렉토리 내의 모든 파일 순회
    for filename in os.listdir(directory_path):
        if filename.endswith(('.xlsx')):
            file_path = os.path.join(directory_path, filename)
            
            try:
                # 엑셀 파일 읽기
                df = pd.read_excel(file_path)
                
                # 시간 관련 컬럼 처리
                time_columns = [col for col in df.columns if 'X [s]' in col]

                # 첫 번째 X [s] 열만 유지하고 나머지는 제거
                columns_to_drop = time_columns[1:]
                df = df.drop(columns=columns_to_drop)
                
                # 처리된 데이터프레임을 딕셔너리에 저장
                excel_files[filename] = df
                
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                continue
    
    return excel_files

#extract ACC : R_IMU ACC, L_IMU ACC
def ACC_extract(df):
    acc_columns = ['X [s]'] + [col for col in df.columns if 'ACC.Z' in col and '[g]' in col]
    return df[acc_columns]

#extract GYRO : R_IMU GYRO, L_IMU GYRO
def GYRO_extract(df):
    gyro_columns = ['X [s]'] + [col for col in df.columns if 'GYRO.Z' in col and '[°/s]' in col]
    return df[gyro_columns]

#extract EMG : R BT, R ST, L BF, L ST
def EMG_extract(df):
    emg_columns = ['X [s]'] + [col for col in df.columns if 'EMG' in col]
    return df[emg_columns]

if __name__ == "__main__":
    directory_path = '/Users/kwonsoomin/python_sm/sprint_data' #파일경로 수정!!!!!
    excel_data = read_excel_files(directory_path)
    
    for filename, df in excel_data.items():
        print(f"\n=== Processing {filename} ===")
        acc_data = ACC_extract(df)
        gyro_data = GYRO_extract(df)
        emg_data = EMG_extract(df)
        