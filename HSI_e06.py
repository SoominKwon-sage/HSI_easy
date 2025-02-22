from HSI_e01 import read_excel_files
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

def get_injury_side(data_dir):
    """
    피실험자별 그룹 분류 및 분석할 다리 방향 결정
    Returns:
        dict: {파일명: {'group': 상태, 'side': 데이터사용방향}}
    """
    # 고유한 파일명만 추출 (방향 정보 제외)
    file_list = [f for f in os.listdir(data_dir) if f.endswith('interpolated.xlsx')]
    subject_list = list(set([f.split('_left_')[0].split('_right_')[0] for f in file_list]))
    
    if not subject_list:
        print("No data files found.")
        return {}
    
    subject_data = {}
    print("\n=== 피실험자 구분 입력 ===")
    print(" - Control 그룹: c")
    print(" - Left Injury 그룹: l")
    print(" - Right Injury 그룹: r")
    print(" - 제외: skip")
    
    for subject in sorted(subject_list):
        group = input(f"\n{subject}의 상태 (c/l/r/skip): ").lower()
        while group not in ['c', 'l', 'r', 'skip']:
            print("올바른 상태를 입력해주세요 (c/l/r/skip)")
            group = input(f"{subject}의 상태: ").lower()
        
        if group == 'skip':
            continue
            
        subject_data[subject] = {
            'group': 'control' if group == 'c' else 'injury',
            'side': 'right' if group in ['c', 'r'] else 'left'
        }
    
    return subject_data

def plot_single_group(time_points, data, title, ylabel, output_path):
    """단일 그룹의 데이터를 그래프로 생성"""
    plt.figure(figsize=(10, 6))
    plt.plot(time_points, data['mean'], 'b-', linewidth=2)
    plt.fill_between(time_points, 
                    data['mean'] - data['std'],
                    data['mean'] + data['std'], 
                    color='blue', alpha=0.2)
    
    plt.title(title)
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel(ylabel)
    plt.grid(True, alpha=0.3)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

def plot_single_emg_group(time_points, data, group, output_path):
    """단일 그룹의 EMG 데이터를 그래프로 생성 (BF, ST 함께 표시)"""
    plt.figure(figsize=(10, 6))
    
    # BF 데이터
    plt.plot(time_points, data['BF']['mean'], 'b-', 
            label='BF', linewidth=2)
    plt.fill_between(time_points, 
                    data['BF']['mean'] - data['BF']['std'],
                    data['BF']['mean'] + data['BF']['std'], 
                    color='blue', alpha=0.2)
    
    # ST 데이터
    plt.plot(time_points, data['ST']['mean'], 'r-', 
            label='ST', linewidth=2)
    plt.fill_between(time_points, 
                    data['ST']['mean'] - data['ST']['std'],
                    data['ST']['mean'] + data['ST']['std'], 
                    color='red', alpha=0.2)
    
    plt.title(f'EMG - {group} Group')
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel('EMG (μV)')
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

def save_stats_to_excel(results, output_dir):
    """평균값과 표준편차를 하나의 엑셀 파일에 여러 시트로 저장"""
    os.makedirs(output_dir, exist_ok=True)  # 수정된 부분: output_dir 자동 생성
    time_points = np.linspace(0, 100, 101)
    output_path = os.path.join(output_dir, 'analysis_stats.xlsx')
    
    # ExcelWriter 객체 생성
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # IMU와 ACC 데이터 저장
        for measurement in ['IMU', 'ACC']:
            data = {
                'Time(%)': time_points,
                'Control_Mean': results['control'][measurement]['mean'],
                'Control_STD': results['control'][measurement]['std'],
                'Injury_Mean': results['injury'][measurement]['mean'],
                'Injury_STD': results['injury'][measurement]['std']
            }
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=measurement, index=False)
        
        # EMG 데이터 저장 (BF, ST)
        for muscle in ['BF', 'ST']:
            data = {
                'Time(%)': time_points,
                'Control_Mean': results['control'][muscle]['mean'],
                'Control_STD': results['control'][muscle]['std'],
                'Injury_Mean': results['injury'][muscle]['mean'],
                'Injury_STD': results['injury'][muscle]['std']
            }
            df = pd.DataFrame(data)
            sheet_name = f'EMG_{muscle}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Saved all statistics to: {output_path}")

def analyze_injury_data(interpolated_dir, subject_data, output_dir):
    """부상 데이터 분석 및 시각화"""
    os.makedirs(output_dir, exist_ok=True)  # 수정된 부분: output_dir 자동 생성
    stats = {
        'control': {'BF': [], 'ST': [], 'IMU': [], 'ACC': []},
        'injury': {'BF': [], 'ST': [], 'IMU': [], 'ACC': []}
    }
    
    # 데이터 수집 및 통계 계산
    for subject, info in subject_data.items():
        try:
            filename = f"{subject}_{info['side']}_interpolated.xlsx"
            file_path = os.path.join(interpolated_dir, filename)
            
            # IMU 데이터
            imu_data = pd.read_excel(file_path, sheet_name='IMU_sprint1', index_col=0)
            imu_data.index = imu_data.index.str.rstrip('%').astype(float)
            stats[info['group']]['IMU'].extend(imu_data.values.T)
            
            # ACC 데이터
            acc_data = pd.read_excel(file_path, sheet_name='ACC_sprint1', index_col=0)
            acc_data.index = acc_data.index.str.rstrip('%').astype(float)
            stats[info['group']]['ACC'].extend(acc_data.values.T)
            
            # EMG 데이터 (BF, ST)
            for muscle in ['BF', 'ST']:
                sheet_name = f'{muscle}_sprint1'
                emg_data = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)
                emg_data.index = emg_data.index.str.rstrip('%').astype(float)
                stats[info['group']][muscle].extend(emg_data.values.T)
            
            print(f"Successfully processed: {filename}")
            
        except Exception as e:
            print(f"Error processing {subject}: {str(e)}")
            continue
    
    # 평균과 표준편차 계산
    time_points = np.linspace(0, 100, 101)
    results = {group: {} for group in ['control', 'injury']}
    
    for group in results:
        group_has_data = False
        for key in stats[group]:
            if stats[group][key]:
                group_has_data = True
                data = np.array(stats[group][key])
                results[group][key] = {
                    'mean': np.mean(data, axis=0),
                    'std': np.std(data, axis=0)
                }

        if not group_has_data:
            print(f"No data found for {group} group")
            continue

        # 각 그룹별로 데이터가 있는 경우 그래프 생성
        if 'IMU' in results[group]:
            plot_single_group(time_points, 
                            results[group]['IMU'],
                            f'IMU - {group.capitalize()} Group',
                            'Angular Velocity (°/s)',
                            os.path.join(output_dir, f'IMU_{group}.png'))
        
        if 'ACC' in results[group]:
            plot_single_group(time_points,
                            results[group]['ACC'],
                            f'ACC - {group.capitalize()} Group',
                            'Acceleration (g)',
                            os.path.join(output_dir, f'ACC_{group}.png'))
        
        if all(key in results[group] for key in ['BF', 'ST']):
            plot_single_emg_group(time_points, 
                                results[group],
                                group.capitalize(),
                                os.path.join(output_dir, f'EMG_{group}.png'))
    
    # 통계 데이터를 엑셀로 저장
    save_stats_to_excel(results, output_dir)
    
    return results

if __name__ == "__main__":
    data_dir = '/Users/kwonsoomin/python_sm/sprint_data'  # Raw 데이터 경로는 그대로 유지
    interpolated_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '05_InterpolatedData')  # 수정된 부분
    output_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '06_InjuryAnalysis')         # 수정된 부분

    os.makedirs(interpolated_dir, exist_ok=True)  # 수정된 부분: 폴더 자동 생성
    os.makedirs(output_dir, exist_ok=True)          # 수정된 부분: 폴더 자동 생성

    subject_data = get_injury_side(interpolated_dir)
    if subject_data:
        print("\nAnalyzing data...")
        results = analyze_injury_data(interpolated_dir, subject_data, output_dir)
        print("\nAnalysis completed successfully!")