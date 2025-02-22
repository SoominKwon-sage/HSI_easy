# Data interpolation
from HSI_e01 import read_excel_files, ACC_extract, GYRO_extract, EMG_extract
from HSI_e02 import find_gait_cycles
from HSI_e04 import find_sprint_intervals, find_cycles_in_sprint, select_middle_cycles
import numpy as np
from scipy.interpolate import interp1d
import pandas as pd
import os

def interpolate_cycle_data(time, data, cycle_indices, num_points=101):
    """
    각 사이클의 데이터를 지정된 개수의 포인트로 보간
    
    Parameters:
    - time: 시간 데이터
    - data: 보간할 데이터 (IMU, ACC, 또는 EMG)
    - cycle_indices: 사이클의 시작 인덱스 리스트
    - num_points: 보간할 포인트 개수 (기본값: 101)
    
    Returns:
    - interpolated_cycles: 보간된 데이터 리스트
    """
    interpolated_cycles = []
    
    for i in range(len(cycle_indices)):
        # 마지막 사이클의 경우, 다음 사이클 인덱스 대신 길이를 이용
        start_idx = cycle_indices[i]
        if i < len(cycle_indices) - 1:
            end_idx = cycle_indices[i + 1]
        else:
            # 마지막 사이클의 경우, 이전 사이클과 같은 길이를 사용
            cycle_length = cycle_indices[i] - cycle_indices[i-1]
            end_idx = min(start_idx + cycle_length, len(time)-1)
        
        # 현재 사이클의 시간과 데이터 추출
        cycle_time = time[start_idx:end_idx+1]
        cycle_data = data[start_idx:end_idx+1]
        
        # 데이터 포인트가 너무 적으면 스킵
        if len(cycle_time) < 4:  # cubic interpolation requires at least 4 points
            print(f"Warning: Cycle {i+1} has too few points ({len(cycle_time)}), skipping...")
            continue
        
        # 시간을 0-100%로 정규화
        normalized_time = np.linspace(0, 100, len(cycle_time))
        target_time = np.linspace(0, 100, num_points)
        
        # 스플라인 보간 수행
        try:
            interpolator = interp1d(normalized_time, cycle_data, kind='cubic')
            interpolated_data = interpolator(target_time)
            interpolated_cycles.append(interpolated_data)
        except Exception as e:
            print(f"Warning: Failed to interpolate cycle {i+1}: {str(e)}")
            continue
    
    return interpolated_cycles

def save_interpolated_data(interpolated_data, filename, output_dir=None):
    """
    보간된 데이터를 엑셀 파일로 저장 (오른쪽과 왼쪽 다리 데이터를 별도의 파일로 저장)
    
    Parameters:
    - interpolated_data: 보간된 데이터 딕셔너리
    - filename: 원본 파일 이름
    - output_dir: 저장할 디렉토리 (None인 경우 동적으로 생성)
    """
    try:
        if output_dir is None:
            output_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '05_InterpolatedData')
        os.makedirs(output_dir, exist_ok=True)
        
        base_filename = os.path.splitext(os.path.basename(filename))[0]
    
        # 오른쪽과 왼쪽 다리 데이터를 별도의 파일로 저장
        for side in ['right', 'left']:
            output_filename = os.path.join(output_dir, f"{base_filename}_{side}_interpolated.xlsx")
        
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                for data_type in ['IMU', 'ACC', 'BF', 'ST']:
                    for category, cycles in interpolated_data[side][data_type].items():
                        sheet_name = f"{data_type}_sprint{category}"
                    
                        # 각 사이클의 데이터를 DataFrame으로 변환
                        cycle_data = {}
                        for i, cycle in enumerate(cycles):
                            cycle_data[f'Cycle_{i+1}'] = cycle
                    
                        df = pd.DataFrame(cycle_data)
                        df.index = [f'{i}%' for i in range(101)]
                        df.to_excel(writer, sheet_name=sheet_name)
                    print(f"Saved {side} leg {sheet_name}")
        
        print(f"Interpolation completed successfully: {base_filename}")
        return True

    except Exception as e:
        print(f"Error during interpolation: {str(e)}")
        return False

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
        
        # 데이터 보간을 위한 딕셔너리
        interpolated_data = {
            'right': {'IMU': {}, 'ACC': {}, 'BF': {}, 'ST': {}},
            'left': {'IMU': {}, 'ACC': {}, 'BF': {}, 'ST': {}}
        }
        
        print("\n=== Processing Right Leg Data ===")
        # 오른쪽 다리 데이터 보간
        for category, cycles in right_selected.items():
            if len(cycles) > 1:
                interpolated_data['right']['IMU'][category] = interpolate_cycle_data(time_data, right_gyro, cycles)
                interpolated_data['right']['ACC'][category] = interpolate_cycle_data(time_data, right_acc, cycles)
                interpolated_data['right']['BF'][category] = interpolate_cycle_data(time_data, right_emg['BF'], cycles)
                interpolated_data['right']['ST'][category] = interpolate_cycle_data(time_data, right_emg['ST'], cycles)
        print("\n=== Processing Left Leg Data ===")

        # 왼쪽 다리 데이터 보간
        for category, cycles in left_selected.items():
            if len(cycles) > 1:
                interpolated_data['left']['IMU'][category] = interpolate_cycle_data(time_data, left_gyro, cycles)
                interpolated_data['left']['ACC'][category] = interpolate_cycle_data(time_data, left_acc, cycles)
                interpolated_data['left']['BF'][category] = interpolate_cycle_data(time_data, left_emg['BF'], cycles)
                interpolated_data['left']['ST'][category] = interpolate_cycle_data(time_data, left_emg['ST'], cycles)
        
        # 결과 저장
        save_interpolated_data(interpolated_data, filename)