import sys
import glob
import os

policy_path = sys.argv[1]
policies = glob.glob(policy_path + '\Policy*')
print('共找到 ', len(policies), ' 組資料如下')

for p in policies:
    print(p.split('\\')[-1])

print('開始檢查:')


from baseline import Policy


for index, policy_dir in enumerate(policies):
    print('第', index+1, '組資料:' , policy_dir)

    # --- Check 1: coregistration_Error_file ---
    coregistration_Error_file = os.path.join(policy_dir, 'Report_3coregistration_Error.txt')
    if not os.path.isfile(coregistration_Error_file):
        print(f'    在 {policy_dir} 中找不到 Report_3coregistration_Error.txt, 無法產生誤差分析')
        continue

    # --- Check 2: coherence_file ---
    coherence_file = os.path.join(policy_dir, 'coherence_phase', 'ifg_coh_filt_coh_compare')
    if not os.path.isfile(coherence_file):
        print(f'    在 {policy_dir} 中找不到 coherence_phase\\ifg_coh_filt_coh_compare, 無法產生同調性分析')
        continue

    # --- Check 3: postprocessing_file ---
    postprocessing_file = os.path.join(policy_dir, 'postprocessing', 'shortbaseline')
    if not os.path.isfile(postprocessing_file):
        print(f'    在 {policy_dir} 中找不到 postprocessing\\shortbaseline, 無法分析基線資訊')
        continue

    # --- Check 4: shortbaseline_plot_path ---
    shortbaseline_plot_path = os.path.join(policy_dir, 'shortbaseline_plot.eps')
    if not os.path.isfile(shortbaseline_plot_path):
        print(f'    在 {policy_dir} 中找不到 shortbaseline_plot.eps, 無法貼上大張基線圖')
        continue
    
    try:
        policy = Policy(policy_dir=policy_dir, index=0)
        missing_bmp = policy.check_bmp_path()
        if missing_bmp is not None:
            print('    找不到', missing_bmp, '無法產生基線圖檔')
            continue
    except Exception as e:
        print(f"    在 {policy_dir} 中處理 baseline 時發生錯誤")
        import traceback
        traceback.print_exc()
        continue

    # If all checks passed for this policy:
    print('    路徑正確')

print('\n檢查完畢')
