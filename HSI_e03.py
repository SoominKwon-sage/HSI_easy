#IMU 기준으로 EMG, ACC phase detection
from HSI_e01 import read_excel_files, ACC_extract, GYRO_extract, EMG_extract
from HSI_e02 import find_gait_cycles
import pandas as pd
import os

def extract_peak_data(valleys, time_data, gyro_data, acc_data, emg_data):
    """
    IMU 음의 피크 시점의 모든 센서 데이터 추출
    
    Parameters:
    - valleys: IMU 음의 피크 인덱스
    - time_data: 시간 데이터
    - gyro_data: GYRO 데이터
    - acc_data: ACC 데이터
    - emg_data: EMG 데이터
    
    Returns:
    - peak_data: 피크 시점의 모든 데이터를 포함하는 딕셔너리
    """
    peak_data = {
        'Time': time_data[valleys],
        'IMU_Peak': gyro_data[valleys],
        'ACC': acc_data[valleys],
        'BF': emg_data['BF'].iloc[valleys],
        'ST': emg_data['ST'].iloc[valleys]
    }
    
    return peak_data

def save_peak_data_to_excel(right_data, left_data, filename, output_dir=None):
    """
    피크 데이터를 엑셀 파일로 저장
    """
    try:
        if output_dir is None:
            output_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '03_PeakData')
        os.makedirs(output_dir, exist_ok=True)
        
        # 데이터프레임 생성
        right_df = pd.DataFrame(right_data)
        right_df.index = range(1, len(right_df) + 1)
        right_df.index.name = 'Peak_Number'
        
        left_df = pd.DataFrame(left_data)
        left_df.index = range(1, len(left_df) + 1)
        left_df.index.name = 'Peak_Number'
        
        # 파일명 준비
        base_filename = os.path.splitext(os.path.basename(filename))[0]
        output_filename = os.path.join(output_dir, f"{base_filename}_peak_data.xlsx")
        
        # 엑셀 파일로 저장
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            right_df.to_excel(writer, sheet_name='Right_Leg_Peaks')
            left_df.to_excel(writer, sheet_name='Left_Leg_Peaks')
            
        print(f"\nPeak data saved to: {output_filename}")
        
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
        right_gyro = gyro_data.iloc[:, 1]  # 오른쪽 GYRO
        right_acc = acc_data.iloc[:, 1]    # 오른쪽 ACC
        right_emg = {
            'BF': emg_data.iloc[:, 1],     # 오른쪽 BF
            'ST': emg_data.iloc[:, 2]      # 오른쪽 ST
        }
        
        # 왼쪽 다리 데이터
        left_gyro = gyro_data.iloc[:, 2]   # 왼쪽 GYRO
        left_acc = acc_data.iloc[:, 2]     # 왼쪽 ACC
        left_emg = {
            'BF': emg_data.iloc[:, 3],     # 왼쪽 BF
            'ST': emg_data.iloc[:, 4]      # 왼쪽 ST
        }
        
        # 보행 사이클의 음의 피크 찾기
        right_valleys, _ = find_gait_cycles(time_data, right_gyro, 'Right')
        left_valleys, _ = find_gait_cycles(time_data, left_gyro, 'Left')
        
        # 피크 시점의 데이터 추출
        right_peak_data = extract_peak_data(
            right_valleys, time_data, right_gyro, right_acc, right_emg)
        left_peak_data = extract_peak_data(
            left_valleys, time_data, left_gyro, left_acc, left_emg)
        
        # 엑셀 파일로 저장
        save_peak_data_to_excel(right_peak_data, left_peak_data, filename)