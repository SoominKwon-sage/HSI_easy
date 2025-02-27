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


# ============ [단일 그룹 그래프 함수들] ============

def plot_single_sensor(time_points, sensor_data, title, ylabel, output_path):
    """
    단일 그룹(IMU/ACC)을 그리는 그래프
    sensor_data = {'mean': ..., 'std': ...}
    """
    plt.figure(figsize=(10, 6))
    plt.plot(time_points, sensor_data['mean'], 'r-', linewidth=2, label='Injury')
    plt.fill_between(time_points, 
                     sensor_data['mean'] - sensor_data['std'],
                     sensor_data['mean'] + sensor_data['std'],
                     color='red', alpha=0.2)
    plt.title(title)
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel(ylabel)
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved single-group graph to: {output_path}")


def plot_single_emg(time_points, emg_data, group_name, output_path):
    """
    단일 그룹 EMG(BF, ST) 그래프
    emg_data = {'BF': {'mean':..., 'std':...}, 'ST': {...}}
    """
    plt.figure(figsize=(10, 6))
    plt.plot(time_points, emg_data['BF']['mean'], 'r-', label='BF', linewidth=2)
    plt.fill_between(time_points, 
                     emg_data['BF']['mean'] - emg_data['BF']['std'],
                     emg_data['BF']['mean'] + emg_data['BF']['std'],
                     color='red', alpha=0.2)
    
    plt.plot(time_points, emg_data['ST']['mean'], 'm-', label='ST', linewidth=2)
    plt.fill_between(time_points, 
                     emg_data['ST']['mean'] - emg_data['ST']['std'],
                     emg_data['ST']['mean'] + emg_data['ST']['std'],
                     color='magenta', alpha=0.2)
    
    plt.title(f"EMG - {group_name} Group")
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel('EMG (μV)')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved single-group EMG graph to: {output_path}")


# ============ [비교 그래프 함수들] ============

def plot_group_comparison(time_points, control_data, injury_data, title, ylabel, output_path):
    """
    control_data와 injury_data를 비교하여 하나의 그래프에 표시합니다.
    만약 control_data가 None이면 injury_data만 표시합니다.
    """
    plt.figure(figsize=(10, 6))
    
    if control_data is not None:
        plt.plot(time_points, control_data['mean'], 'b-', linewidth=2, label='Control')
        plt.fill_between(time_points, 
                         control_data['mean'] - control_data['std'],
                         control_data['mean'] + control_data['std'],
                         color='blue', alpha=0.2)
    else:
        print("Control 그룹 데이터가 없습니다. Injury 그룹만 표시합니다.")
    
    if injury_data is not None:
        plt.plot(time_points, injury_data['mean'], 'r-', linewidth=2, label='Injury')
        plt.fill_between(time_points, 
                         injury_data['mean'] - injury_data['std'],
                         injury_data['mean'] + injury_data['std'],
                         color='red', alpha=0.2)
    else:
        print("Injury 그룹 데이터가 없습니다. 그래프를 그릴 데이터가 부족합니다.")
        plt.close()
        return
    
    plt.title(title)
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel(ylabel)
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved comparison graph to: {output_path}")


def plot_emg_group_comparison(time_points, control_data, injury_data, group_name, output_path):
    """
    EMG 데이터를 control과 injury 그룹으로 비교하여 하나의 그래프로 생성합니다.
    control_data, injury_data는 {'BF': {'mean':..., 'std':...}, 'ST': {...}} 형태입니다.
    """
    plt.figure(figsize=(10, 6))
    
    if control_data is not None:
        plt.plot(time_points, control_data['BF']['mean'], 'b-', label='BF Control', linewidth=2)
        plt.fill_between(time_points, 
                         control_data['BF']['mean'] - control_data['BF']['std'],
                         control_data['BF']['mean'] + control_data['BF']['std'],
                         color='blue', alpha=0.2)
        
        plt.plot(time_points, control_data['ST']['mean'], 'g-', label='ST Control', linewidth=2)
        plt.fill_between(time_points, 
                         control_data['ST']['mean'] - control_data['ST']['std'],
                         control_data['ST']['mean'] + control_data['ST']['std'],
                         color='green', alpha=0.2)
    else:
        print("Control 그룹 EMG 데이터가 없습니다. Injury 그룹만 표시합니다.")
    
    if injury_data is not None:
        plt.plot(time_points, injury_data['BF']['mean'], 'r-', label='BF Injury', linewidth=2)
        plt.fill_between(time_points, 
                         injury_data['BF']['mean'] - injury_data['BF']['std'],
                         injury_data['BF']['mean'] + injury_data['BF']['std'],
                         color='red', alpha=0.2)
        
        plt.plot(time_points, injury_data['ST']['mean'], 'm-', label='ST Injury', linewidth=2)
        plt.fill_between(time_points, 
                         injury_data['ST']['mean'] - injury_data['ST']['std'],
                         injury_data['ST']['mean'] + injury_data['ST']['std'],
                         color='magenta', alpha=0.2)
    else:
        print("Injury 그룹 EMG 데이터가 없습니다.")
        plt.close()
        return
    
    plt.title(f'EMG - {group_name} Group Comparison')
    plt.xlabel('Gait Cycle (%)')
    plt.ylabel('EMG (μV)')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved EMG comparison graph to: {output_path}")


def save_stats_to_excel(results, output_dir):
    """
    평균값과 표준편차를 하나의 엑셀 파일에 여러 시트로 저장 (control 그룹이 없으면 injury 데이터만 저장)
    만약 저장할 데이터가 없으면 기본 메시지가 담긴 시트를 생성합니다.
    """
    os.makedirs(output_dir, exist_ok=True)
    time_points = np.linspace(0, 100, 101)
    output_path = os.path.join(output_dir, 'analysis_stats.xlsx')
    
    sheet_count = 0  # 생성한 시트 수 추적
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # IMU와 ACC 데이터 저장
        for measurement in ['IMU', 'ACC']:
            data = {'Time(%)': time_points}
            if results.get('control') is not None and measurement in results['control']:
                data['Control_Mean'] = results['control'][measurement]['mean']
                data['Control_STD'] = results['control'][measurement]['std']
            if results.get('injury') is not None and measurement in results['injury']:
                data['Injury_Mean'] = results['injury'][measurement]['mean']
                data['Injury_STD'] = results['injury'][measurement]['std']
            
            if len(data) > 1:
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=measurement, index=False)
                sheet_count += 1
            else:
                print(f"No data for {measurement} found. Skipping sheet creation.")
        
        # EMG 데이터 저장 (BF, ST)
        for muscle in ['BF', 'ST']:
            data = {'Time(%)': time_points}
            if results.get('control') is not None and muscle in results['control']:
                data['Control_Mean'] = results['control'][muscle]['mean']
                data['Control_STD'] = results['control'][muscle]['std']
            if results.get('injury') is not None and muscle in results['injury']:
                data['Injury_Mean'] = results['injury'][muscle]['mean']
                data['Injury_STD'] = results['injury'][muscle]['std']
            
            if len(data) > 1:
                df = pd.DataFrame(data)
                sheet_name = f'EMG_{muscle}'
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheet_count += 1
            else:
                print(f"No data for EMG {muscle} found. Skipping sheet creation.")
        
        # 만약 생성된 시트가 하나도 없다면, 더미 시트를 생성하여 오류를 방지
        if sheet_count == 0:
            dummy_df = pd.DataFrame({"Message": ["No statistical data available."]})
            dummy_df.to_excel(writer, sheet_name="No_Data", index=False)
            print("No statistical data found for any group. Created a dummy sheet.")
    
    print(f"Saved all statistics to: {output_path}")


def analyze_injury_data(interpolated_dir, subject_data, output_dir):
    """부상 데이터 분석 및 시각화 (조건부로 비교 그래프 호출)"""
    os.makedirs(output_dir, exist_ok=True)
    stats = {
        'control': {'BF': [], 'ST': [], 'IMU': [], 'ACC': []},
        'injury': {'BF': [], 'ST': [], 'IMU': [], 'ACC': []}
    }
    
    # 1) 데이터 수집 및 통계 계산
    for subject, info in subject_data.items():
        try:
            filename = f"{subject}_{info['side']}_interpolated.xlsx"
            file_path = os.path.join(interpolated_dir, filename)
            
            # IMU
            imu_data = pd.read_excel(file_path, sheet_name='IMU_sprint1', index_col=0)
            imu_data.index = imu_data.index.str.rstrip('%').astype(float)
            stats[info['group']]['IMU'].extend(imu_data.values.T)
            
            # ACC
            acc_data = pd.read_excel(file_path, sheet_name='ACC_sprint1', index_col=0)
            acc_data.index = acc_data.index.str.rstrip('%').astype(float)
            stats[info['group']]['ACC'].extend(acc_data.values.T)
            
            # EMG (BF, ST)
            for muscle in ['BF', 'ST']:
                sheet_name = f'{muscle}_sprint1'
                emg_data = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)
                emg_data.index = emg_data.index.str.rstrip('%').astype(float)
                stats[info['group']][muscle].extend(emg_data.values.T)
            
            print(f"Successfully processed: {filename}")
            
        except Exception as e:
            print(f"Error processing {subject}: {str(e)}")
            continue
    
    # 2) 평균과 표준편차 계산
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
            results[group] = None
    
    # 3) 조건부 그래프 생성
    control_imu = results['control']['IMU'] if results['control'] and 'IMU' in results['control'] else None
    control_acc = results['control']['ACC'] if results['control'] and 'ACC' in results['control'] else None
    control_bf  = results['control']['BF']  if results['control'] and 'BF'  in results['control'] else None
    control_st  = results['control']['ST']  if results['control'] and 'ST'  in results['control'] else None
    
    injury_imu = results['injury']['IMU'] if results['injury'] and 'IMU' in results['injury'] else None
    injury_acc = results['injury']['ACC'] if results['injury'] and 'ACC' in results['injury'] else None
    injury_bf  = results['injury']['BF']  if results['injury'] and 'BF'  in results['injury'] else None
    injury_st  = results['injury']['ST']  if results['injury'] and 'ST'  in results['injury'] else None
    
    # === IMU ===
    if control_imu and injury_imu:
        # Control과 Injury 모두 존재 -> 비교 그래프
        plot_group_comparison(time_points,
                              control_imu,
                              injury_imu,
                              "IMU Group Comparison",
                              "Angular Velocity (°/s)",
                              os.path.join(output_dir, "IMU_comparison.png"))
    elif injury_imu:
        # Control이 없고 Injury만 존재 -> 단일 그래프
        plot_single_sensor(time_points,
                           injury_imu,
                           "IMU - Injury Group",
                           "Angular Velocity (°/s)",
                           os.path.join(output_dir, "IMU_injury.png"))
    
    # === ACC ===
    if control_acc and injury_acc:
        plot_group_comparison(time_points,
                              control_acc,
                              injury_acc,
                              "ACC Group Comparison",
                              "Acceleration (g)",
                              os.path.join(output_dir, "ACC_comparison.png"))
    elif injury_acc:
        plot_single_sensor(time_points,
                           injury_acc,
                           "ACC - Injury Group",
                           "Acceleration (g)",
                           os.path.join(output_dir, "ACC_injury.png"))
    
    # === EMG ===
    if control_bf and control_st and injury_bf and injury_st:
        # Control & Injury 모두 BF, ST 데이터 있음 -> 비교 그래프
        control_emg_data = {'BF': control_bf, 'ST': control_st}
        injury_emg_data  = {'BF': injury_bf,  'ST': injury_st}
        plot_emg_group_comparison(time_points,
                                  control_emg_data,
                                  injury_emg_data,
                                  "EMG",
                                  os.path.join(output_dir, "EMG_comparison.png"))
    elif injury_bf and injury_st:
        # Control이 없거나 데이터 부족, Injury만 있음 -> 단일 그래프
        single_emg_data = {'BF': injury_bf, 'ST': injury_st}
        plot_single_emg(time_points,
                        single_emg_data,
                        "Injury",
                        os.path.join(output_dir, "EMG_injury.png"))
    
    # 4) 통계 데이터를 엑셀로 저장
    save_stats_to_excel(results, output_dir)
    
    return results


if __name__ == "__main__":
    data_dir = '/Users/kwonsoomin/python_sm/sprint_data'
    interpolated_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '05_InterpolatedData')
    output_dir = os.path.join(os.getcwd(), 'HSI_DataProcessing', '06_InjuryAnalysis')
    
    os.makedirs(interpolated_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    subject_data = get_injury_side(interpolated_dir)
    if subject_data:
        print("\nAnalyzing data...")
        results = analyze_injury_data(interpolated_dir, subject_data, output_dir)
        print("\nAnalysis completed successfully!")
