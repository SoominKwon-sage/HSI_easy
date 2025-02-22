# IMU phase detection
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from HSI_e01 import read_excel_files, GYRO_extract

def find_gait_cycles(time, gyro_data, side='Right', min_distance=20, window_size=50, peak_threshold=-200):
    """
    보행 사이클 찾기 함수
    Parameters:
    - time: 시간 데이터
    - gyro_data: 자이로스코프 데이터
    - side: 'Right' 또는 'Left'
    - min_distance: 피크 간 최소 거리 (샘플 수)
    - window_size: 로컬 최소값을 찾을 윈도우 크기
    - peak_threshold: 이 값보다 작은 음의 피크만 유효한 보행 사이클로 간주
    
    Returns:
    - valleys: 음의 피크 위치
    - cycles: 각 사이클 정보
    """
    valleys = []
    i = window_size
    
    while i < len(gyro_data) - window_size:
        window = gyro_data[i-window_size:i+window_size]
        local_min_idx = i - window_size + np.argmin(window)
        
        if local_min_idx == i and gyro_data[i] < peak_threshold:
            if not valleys or (local_min_idx - valleys[-1]) >= min_distance:
                valleys.append(local_min_idx)
                i = local_min_idx + min_distance
            else:
                if gyro_data[local_min_idx] < gyro_data[valleys[-1]]:
                    valleys[-1] = local_min_idx
                i += 1
        else:
            i += 1
    
    valleys = np.array(valleys)
    
    cycles = []
    for i in range(len(valleys)-1):
        cycle_start = time[valleys[i]]
        cycle_end = time[valleys[i+1]]
        cycle_duration = cycle_end - cycle_start
        
        peak_velocity = gyro_data[valleys[i]]
        cycle_data = gyro_data[valleys[i]:valleys[i+1]]
        max_velocity = np.max(np.abs(cycle_data))
        
        cycles.append({
            'start_time': cycle_start,
            'end_time': cycle_end,
            'duration': cycle_duration,
            'start_idx': valleys[i],
            'end_idx': valleys[i+1],
            'max_velocity': max_velocity,
            'peak_velocity': peak_velocity
        })
    
    return valleys, cycles

def plot_gait_cycles(time_data, gyro_data, valleys, filename, side='Right', peak_threshold=-200):
    """보행 사이클 시각화 함수"""
    plt.plot(time_data, gyro_data, label=f'{side} Gyro', alpha=0.8)
    plt.plot(time_data[valleys], gyro_data[valleys], 
             'rx', label='Gait Cycle (peak < -200°/s)', markersize=10)
    
    for i, valley in enumerate(valleys):
        plt.annotate(f'{i+1}', (time_data[valley], gyro_data[valley]),
                    xytext=(5, 10), textcoords='offset points', fontsize=8)
    
    plt.title(f'{filename} - {side} Leg Gait Cycles (Excluding Rest Periods)')
    plt.xlabel('Time (s)')
    plt.ylabel('Angular Velocity (°/s)')
    plt.axhline(y=peak_threshold, color='r', linestyle='--', alpha=0.3, 
                label=f'Threshold ({peak_threshold}°/s)')
    plt.legend()
    plt.grid(True)

if __name__ == "__main__":
    # 파일 읽기
    directory_path = '/Users/kwonsoomin/python_sm/sprint_data'
    excel_data = read_excel_files(directory_path)
    
    for filename, df in excel_data.items():
        print(f"\n=== Processing {filename} ===")
        # GYRO 데이터 추출
        gyro_data = GYRO_extract(df)
        
        # 시간 데이터와 GYRO 데이터 분리
        time_data = gyro_data['X [s]']
        right_gyro = gyro_data.iloc[:, 1]  # 오른쪽 GYRO 데이터
        left_gyro = gyro_data.iloc[:, 2]   # 왼쪽 GYRO 데이터
        
        # 보행 사이클 분석
        right_valleys, right_cycles = find_gait_cycles(time_data, right_gyro, 'Right')
        left_valleys, left_cycles = find_gait_cycles(time_data, left_gyro, 'Left')
        
        # 그래프 그리기
        plt.figure(figsize=(15, 10))
        
        plt.subplot(2, 1, 1)
        plot_gait_cycles(time_data, right_gyro, right_valleys, filename, 'Right')
        
        plt.subplot(2, 1, 2)
        plot_gait_cycles(time_data, left_gyro, left_valleys, filename, 'Left')
        
        plt.tight_layout()
        plt.show()
        
