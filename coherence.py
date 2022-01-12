import os
import numpy as np
from numpy.polynomial.polynomial import polyfit

import matplotlib.pyplot as plt
from matplotlib import rcParams
plt.rcParams["figure.figsize"] = (6,6)

#以下設定Times New Roman字體
#plt.rcParams["font.family"] = "Times New Roman"
#以下處理標楷體字體
plt.rcParams['font.sans-serif'] = ['DFKai-SB']#['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False



def extract_data(coherence_file):
    before = []
    after = []
    with open(coherence_file) as f:
        for index, line in enumerate(f.readlines()):
            before.append(float(line.split('\t')[-2].strip()))
            after.append(float(line.split('\t')[-1].strip()))
 
    return before, after        


def export_result(coherence_file, policy_path, index):
    
    before, after = extract_data(coherence_file)
     
    plt.scatter(before, after, facecolors='none', edgecolors='b')
    plt.xlabel("濾波前同調性", fontsize=18)
    plt.ylabel("濾波後同調性", fontsize=18)
    plt.title("濾波前後同調性比較圖 (" + str(index) + ')', fontsize=20)
    plt.xlim(0, 1)
    plt.xticks([0, 0.5, 1])
    plt.minorticks_on()
    plt.ylim(0, 1)
    plt.yticks([0, 0.5, 1])
    plt.grid(which='minor',alpha=0.3)
    plt.grid(which='major',alpha=0.6, linewidth=1.5)
    
    b, m = polyfit(before, after,1)
    ggg = [min(before), max(before)]
    plt.plot(ggg, b + m * np.array(ggg), color='black', linestyle='--', dashes=(5, 5))

    #plt.show()
    img_path = policy_path + '\\coherence.png'
    plt.savefig(img_path, dpi=100)
    os.system('start ' + img_path)
    plt.clf()




if __name__ == '__main__':

    import sys
    policy_path = sys.argv[1]
    print('=== 同調性報表產生器 ===')
    print('啟動中...')
    print('正在 ' + policy_path + ' 位置下尋找 Policy 資料夾... ', end='')

    import glob
    policies = glob.glob(policy_path + '\Policy*')
    print('共找到 ', len(policies), ' 組資料如下')

    policies = [p.split('\\')[-1] for p in policies]
    for p in policies:
        print(p)

    print('開始處理:')

    import os
    
    for index, policy_dir in enumerate(policies):
        # 如果沒有 ifg_coh_filt_coh_compare 檔案的話，這筆 Policy 會被跳過

        coherence_file = policy_path + '\\' + policy_dir + '\\coherence_phase\\ifg_coh_filt_coh_compare'
        if not os.path.isfile(coherence_file):
            print('找不到 ifg_coh_filt_coh_compare, 跳過', policy_dir)
            continue
        
        print('開始處理', policy_dir, '...')#, end = '')

        export_result(coherence_file, policy_path + '\\' + policy_dir, index+1)