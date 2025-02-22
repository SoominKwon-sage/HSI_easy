# Main script
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from HSI_e01 import read_excel_files, ACC_extract, GYRO_extract, EMG_extract
from HSI_e02 import find_gait_cycles, plot_gait_cycles
from HSI_e03 import extract_peak_data, save_peak_data_to_excel
from HSI_e04 import (find_sprint_intervals, find_cycles_in_sprint, 
                   select_middle_cycles, save_interval_data_to_excel)
from HSI_e05 import interpolate_cycle_data, save_interpolated_data
from HSI_e06 import get_injury_side, analyze_injury_data

def create_directories(base_dir):
    """분석 결과를 저장할 디렉토리 생성"""
    directories = {
        'peak_data': os.path.join(base_dir, 'HSI_DataProcessing', '03_PeakData'),
        'interval_data': os.path.join(base_dir, 'HSI_DataProcessing', '04_IntervalData'),
        'interpolated_data': os.path.join(base_dir, 'HSI_DataProcessing', '05_InterpolatedData'),
        'injury_analysis': os.path.join(base_dir, 'HSI_DataProcessing', '06_InjuryAnalysis')
    }
    
    for directory in directories.values():
        os.makedirs(directory, exist_ok=True)
    
    return directories

def process_file(filename, df, directories):
    """각 파일에 대한 처리 과정"""
    print(f"\n{'='*20} Processing {filename} {'='*20}")
    
    # 1. Data Extraction
    print("\n[Step 1] Extracting sensor data...")
    gyro_data = GYRO_extract(df)
    acc_data = ACC_extract(df)
    emg_data = EMG_extract(df)
    
    time_data = gyro_data['X [s]']
    
    # Right leg data
    right_gyro = gyro_data.iloc[:, 1]
    right_acc = acc_data.iloc[:, 1]
    right_emg = {
        'BF': emg_data.iloc[:, 1],
        'ST': emg_data.iloc[:, 2]
    }
    
    # Left leg data
    left_gyro = gyro_data.iloc[:, 2]
    left_acc = acc_data.iloc[:, 2]
    left_emg = {
        'BF': emg_data.iloc[:, 3],
        'ST': emg_data.iloc[:, 4]
    }
    
    # 2. Gait Cycle Analysis
    print("\n[Step 2] Analyzing gait cycles...")
    right_valleys, right_cycles = find_gait_cycles(time_data, right_gyro, 'Right')
    left_valleys, left_cycles = find_gait_cycles(time_data, left_gyro, 'Left')
    
    # 3. Peak Data Analysis
    print("\n[Step 3] Extracting peak data...")
    right_peak_data = extract_peak_data(
        right_valleys, time_data, right_gyro, right_acc, right_emg)
    left_peak_data = extract_peak_data(
        left_valleys, time_data, left_gyro, left_acc, left_emg)
    save_peak_data_to_excel(right_peak_data, left_peak_data, filename, 
                           directories['peak_data'])
    
    # 4. Sprint Interval Analysis
    print("\n[Step 4] Analyzing sprint intervals...")
    right_intervals = find_sprint_intervals(right_gyro.values, time_data.values)
    left_intervals = find_sprint_intervals(left_gyro.values, time_data.values)
    
    right_categorized = find_cycles_in_sprint(right_valleys, right_intervals)
    left_categorized = find_cycles_in_sprint(left_valleys, left_intervals)
    
    right_selected = select_middle_cycles(right_categorized)
    left_selected = select_middle_cycles(left_categorized)
    
    save_interval_data_to_excel(
        right_selected, left_selected, time_data,
        right_gyro, left_gyro, right_acc, left_acc,
        right_emg, left_emg, filename, directories['interval_data']
    )
    
    # 5. Data Interpolation
    print("\n[Step 5] Interpolating cycle data...")
    interpolated_data = {
        'right': {'IMU': {}, 'ACC': {}, 'BF': {}, 'ST': {}},
        'left': {'IMU': {}, 'ACC': {}, 'BF': {}, 'ST': {}}
    }
    
    # Right leg interpolation
    for category, cycles in right_selected.items():
        if len(cycles) > 1:
            interpolated_data['right']['IMU'][category] = interpolate_cycle_data(
                time_data, right_gyro, cycles)
            interpolated_data['right']['ACC'][category] = interpolate_cycle_data(
                time_data, right_acc, cycles)
            interpolated_data['right']['BF'][category] = interpolate_cycle_data(
                time_data, right_emg['BF'], cycles)
            interpolated_data['right']['ST'][category] = interpolate_cycle_data(
                time_data, right_emg['ST'], cycles)
    
    # Left leg interpolation
    for category, cycles in left_selected.items():
        if len(cycles) > 1:
            interpolated_data['left']['IMU'][category] = interpolate_cycle_data(
                time_data, left_gyro, cycles)
            interpolated_data['left']['ACC'][category] = interpolate_cycle_data(
                time_data, left_acc, cycles)
            interpolated_data['left']['BF'][category] = interpolate_cycle_data(
                time_data, left_emg['BF'], cycles)
            interpolated_data['left']['ST'][category] = interpolate_cycle_data(
                time_data, left_emg['ST'], cycles)
    
    save_interpolated_data(interpolated_data, filename, directories['interpolated_data'])

def main():
    # Base directory setup: raw 데이터 경로 외에는 동적으로 현재 작업 디렉토리를 기준으로 사용
    base_dir = os.getcwd()
    data_dir = os.path.join(base_dir, 'sprint_data')
    
    print("\n=== HSI Data Analysis Pipeline ===")
    
    # Create necessary directories
    directories = create_directories(base_dir)
    
    # Phase 1: Process each file (Step 1-5)
    print("\n[Phase 1] Reading and processing files...")
    excel_data = read_excel_files(data_dir)
    
    for filename, df in excel_data.items():
        process_file(filename, df, directories)
    
        # Phase 2: Injury Analysis (Step 6)
    print("\n[Phase 2] Performing injury analysis...")
    subject_data = get_injury_side(directories['interpolated_data'])
    
    if subject_data:
        print("\nAnalyzing data and generating visualizations...")
        results = analyze_injury_data(
            directories['interpolated_data'],
            subject_data,
            directories['injury_analysis']
        )
        print("\nAnalysis results have been saved to:")
        print(f"- Graphs: {directories['injury_analysis']}/*.png")
        print(f"- Statistics: {directories['injury_analysis']}/analysis_stats.xlsx")
    
    print("\n=== Analysis Pipeline Completed Successfully! ===")

if __name__ == "__main__":
    main()