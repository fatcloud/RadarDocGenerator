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
    good = True

    # 如果沒有 Report_3coregistration_Error 檔案的話，這筆 Policy 會被跳過
    print('第', index+1, '組資料:' , policy_dir)

    coregistration_Error_file = policy_dir + '\\Report_3coregistration_Error.txt'
    if not os.path.isfile(coregistration_Error_file):
        print('    找不到 Report_3coregistration_Error.txt, 無法產生誤差分析')
        good = False

    coherence_file = policy_dir + '\\coherence_phase\\ifg_coh_filt_coh_compare'
    if not os.path.isfile(coherence_file):
        print('    找不到 ifg_coh_filt_coh_compare, 無法產生同調性分析')
        good = False

    postprocessing_file = policy_dir + '\\postprocessing\\shortbaseline'
    if not os.path.isfile(postprocessing_file):
        print('    找不到', postprocessing_file, '無法分析基線資訊')
        good = False

    shortbaseline_plot_path =policy_dir + '\\shortbaseline_plot.png'
    if not os.path.isfile(shortbaseline_plot_path):
        print('    找不到', shortbaseline_plot_path, '無法貼上大張基線圖')
        good = False

    policy = Policy(policy_dir=policy_dir, index=0)
    missing_bmp = policy.check_bmp_path()
    if missing_bmp is not None:
        print('    找不到', missing_bmp, '無法產生基線圖檔')
        good = False

    if good:
        print('路徑正確')