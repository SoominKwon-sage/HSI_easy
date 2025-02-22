# interval 과 휴식시간 구분
# 필요한 10개의 사이클 찾기

from HSI_e01 import read_excel_files, ACC_extract, GYRO_extract, EMG_extract
from HSI_e02 import find_gait_cycles
import numpy as np
import pandas as pd
import os

def find_sprint_intervals(gyro_data, time_data, velocity_threshold=150, min_rest_duration=3):
    """
    휴식 구간을 기준으로 스프린트 인터벌을 구분
    
    Parameters:
    - gyro_data: GYRO 데이터
    - time_data: 시간 데이터
    - velocity_threshold: 휴식 구간으로 판단할 속도 임계값
    - min_rest_duration: 최소 휴식 시간 (초)
    
    Returns:
    - intervals: {인터벌 번호: [시작 인덱스, 끝 인덱스]} 형태의 딕셔너리
    """
    intervals = {}
    rest_start = None
    interval_count = 1
    current_interval_start = 0
    
    for i in range(len(gyro_data)):
        if abs(gyro_data[i]) <= velocity_threshold:
            if rest_start is None:
                rest_start = i
            elif (time_data[i] - time_data[rest_start]) >= min_rest_duration:
                if current_interval_start < rest_start:
                    intervals[interval_count] = [current_interval_start, rest_start]
                    interval_count += 1
                current_interval_start = i
        else:
            rest_start = None
    
    if current_interval_start < len(gyro_data)-1:
        intervals[interval_count] = [current_interval_start, len(gyro_data)-1]
        
    return intervals

def find_cycles_in_sprint(valleys, interval_bounds):
    """
    각 스프린트 인터벌 내의 사이클들을 찾아서 카테고리화
    
    Parameters:
    - valleys: IMU 음의 피크 위치
    - interval_bounds: 인터벌 경계
    
    Returns:
    - categorized_cycles: {인터벌 번호: [해당 인터벌의 사이클 인덱스]} 형태의 딕셔너리
    """
    categorized_cycles = {}
    
    for category, (start_idx, end_idx) in interval_bounds.items():
        category_cycles = [v for v in valleys if start_idx <= v <= end_idx]
        categorized_cycles[category] = category_cycles
    
    return categorized_cycles

def select_middle_cycles(categorized_cycles, n_cycles=10):
    """
    각 카테고리 내에서 중간 n개의 사이클 선택
    
    Parameters:
    - categorized_cycles: 카테고리별로 분류된 사이클들
    - n_cycles: 선택할 사이클 개수
    
    Returns:
    - selected: {카테고리: [선택된 사이클 인덱스]} 형태의 딕셔너리
    """
    selected = {}
    
    for category, cycles in categorized_cycles.items():
        if len(cycles) < n_cycles:
            continue
            
        mid_point = len(cycles) // 2
        start_idx = mid_point - (n_cycles // 2)
        end_idx = start_idx + n_cycles
        
        if start_idx < 0:
            start_idx = 0
            end_idx = n_cycles
        elif end_idx > len(cycles):
            end_idx = len(cycles)
            start_idx = end_idx - n_cycles
            
        selected[category] = cycles[start_idx:end_idx]
    
    return selected

def save_interval_data_to_excel(right_selected, left_selected, time_data, 
                                right_gyro, left_gyro, right_acc, left_acc,
                                right_emg, left_emg, filename, 
                                output_dir=None):
    """
    선택된 인터벌의 모든 센서 데이터를 엑셀 파일로 저장 (오른쪽/왼쪽 다리 별도 시트)
    """
    try:
        if output_dir is None:
            output_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '04_IntervalData')
        os.makedirs(output_dir, exist_ok=True)
        
        # 오른쪽 다리 데이터 준비
        right_data = []
        for category, cycles in right_selected.items():
            for cycle_idx in cycles:
                right_data.append({
                    'Category': category,
                    'Time': time_data[cycle_idx],
                    'IMU_Peak': right_gyro[cycle_idx],
                    'ACC': right_acc[cycle_idx],
                    'EMG_BF': right_emg['BF'].iloc[cycle_idx],
                    'EMG_ST': right_emg['ST'].iloc[cycle_idx]
                })
        
        # 왼쪽 다리 데이터 준비
        left_data = []
        for category, cycles in left_selected.items():
            for cycle_idx in cycles:
                left_data.append({
                    'Category': category,
                    'Time': time_data[cycle_idx],
                    'IMU_Peak': left_gyro[cycle_idx],
                    'ACC': left_acc[cycle_idx],
                    'EMG_BF': left_emg['BF'].iloc[cycle_idx],
                    'EMG_ST': left_emg['ST'].iloc[cycle_idx]
                })
        
        # DataFrame 생성
        right_df = pd.DataFrame(right_data)
        left_df = pd.DataFrame(left_data)
        
        # 파일명 준비
        base_filename = os.path.splitext(os.path.basename(filename))[0]
        output_filename = os.path.join(output_dir, f"{base_filename}_interval_data.xlsx")
        
        # 엑셀 파일로 저장
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            right_df.to_excel(writer, sheet_name='Right_Leg', index=False)
            left_df.to_excel(writer, sheet_name='Left_Leg', index=False)
        
        print(f"\nInterval data saved to: {output_filename}")
        
    except Exception as e:
        print(f"Error saving to excel: {str(e)}")

if __name__ == "__main__":
    # 파일 읽기
    directory_path = '/Users/kwonsoomin/python_sm/sprint_data'
    excel_data = read_excel_files(directory_path)
    
    for filename, df in excel_data.items():
        print(f"\n=== Processing {filename} ===")
        
        # 데이터 추출
        gyro_data = GYRO_extract(df)
        acc_data = ACC_extract(df)
        emg_data = EMG_extract(df)
        
        # 시간 데이터와 센서 데이터 분리
        time_data = gyro_data['X [s]']
        
        # 오른쪽 다리 데이터
        right_gyro = gyro_data.iloc[:, 1]
        right_acc = acc_data.iloc[:, 1]
        right_emg = {
            'BF': emg_data.iloc[:, 1],
            'ST': emg_data.iloc[:, 2]
        }
        
        # 왼쪽 다리 데이터
        left_gyro = gyro_data.iloc[:, 2]
        left_acc = acc_data.iloc[:, 2]
        left_emg = {
            'BF': emg_data.iloc[:, 3],
            'ST': emg_data.iloc[:, 4]
        }
        
        # IMU 음의 피크 찾기
        right_valleys, _ = find_gait_cycles(time_data, right_gyro, 'Right')
        left_valleys, _ = find_gait_cycles(time_data, left_gyro, 'Left')
        
        # 스프린트 인터벌 찾기
        right_intervals = find_sprint_intervals(right_gyro.values, time_data.values)
        left_intervals = find_sprint_intervals(left_gyro.values, time_data.values)
        
        # 각 인터벌 내의 사이클 찾기
        right_categorized = find_cycles_in_sprint(right_valleys, right_intervals)
        left_categorized = find_cycles_in_sprint(left_valleys, left_intervals)
        
        # 각 카테고리에서 중간 10개 선택
        right_selected = select_middle_cycles(right_categorized)
        left_selected = select_middle_cycles(left_categorized)
        
        # 결과 저장
        save_interval_data_to_excel(right_selected, left_selected, time_data,
                                  right_gyro, left_gyro, 
                                  right_acc, left_acc,
                                  right_emg, left_emg, filename)




