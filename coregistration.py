import os
import numpy as np

import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.patches as patches

#以下設定Times New Roman字體
#plt.rcParams["font.family"] = "Times New Roman"
#以下處理標楷體字體
plt.rcParams['font.sans-serif'] = ['DFKai-SB']#['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False



def extract_data(coregistration_Error_file):
    first_group = []
    second_group = []
    with open(coregistration_Error_file) as f:
        for index, line in enumerate(f.readlines()):
            time = line.split('./TCP_TW_4/M_Ss/')[-1].split('/')[0].strip()
            rrange = line.split('range:')[-1].split('azimuth:')[0].strip()
            azimuth = line.split('azimuth:')[-1].strip()
            #print(time, rrange, azimuth)
            if index % 2 == 0:
                second_group.append([time, rrange, azimuth])
            else:
                first_group.append([time, rrange, azimuth])

    return first_group, second_group


def export_chart(coregistration_Error_file, policy_path):
    
    def min_max_mean_std(data):
        return [str(round(x,2)) for x in [min(data), max(data), np.mean(data), np.std(data)]]

    first_group, second_group = extract_data(coregistration_Error_file)

    G1_range = [float(x[1]) for x in first_group]
    G2_range = [float(x[1]) for x in second_group]

    G1_azimuth = [float(x[2]) for x in first_group]
    G2_azimuth = [float(x[2]) for x in second_group]

    time_periods = [x[0] for x in first_group] 

    def correct_data(title, G1, G2, verbose=True):
        if verbose:
            print('   修正 {title}...'.format(title=title))
        for idx, err in enumerate(G2):
            if err > G1[idx] and err > 0.2:
                if verbose:
                    print('      ' + time_periods[idx] + '  {g2} ===> {g1}'.format(g2=err, g1=G1[idx]))
                G2[idx] = G1[idx]

    correct_data('Range', G1_range, G2_range)
    correct_data('Azimuth', G1_azimuth, G2_azimuth)            

    fig, axs = plt.subplots(2, 1)

    plt.setp(axs, yticks=np.arange(0, 2.2, 0.2))
    
    def plot_something(index, G1, G2, title):
        ax = axs[index]
        ax.set_title(title, fontsize=20, fontweight="bold")
        ax.plot(time_periods, G1, label="第一次配準誤差")
        ax.plot(time_periods, G2, label="第二次配準誤差")
        ax.legend(loc='upper left')
        ax.set_xticklabels(time_periods, rotation=30, ha='right')
        ax.set_ylim(0, 2)
        ax.set_ylabel('誤差值(Pixel)', fontsize=18, fontweight="bold")
        ax.legend(fontsize=10)
        x_max = time_periods[-1]
        ax.set_xlim([0, x_max])
        # 加入偽表格
        chart_items = [
            ['', 'min', 'max', 'mean', 'std'],
            ['第一次配準誤差            '] + min_max_mean_std(G1),
            ['第二次配準誤差            '] + min_max_mean_std(G2)
        ]

        # 底色
        ax.add_patch(
         patches.Rectangle(
            (0.285, 0.64),
            0.34,
            0.32,
            transform=ax.transAxes,
            facecolor = '#F0F0F0',
            fill=True
         ) )

        # 格線
        ax.hlines([0.76, 0.86], 0.2, 0.7, transform=ax.transAxes, linewidth=1, color='white')
        ax.vlines([0.42,0.475, 0.525, 0.575], 0.61, 0.99, transform=ax.transAxes, linewidth=1, color='white')

        for row in range(3):
            for col in range(5):
                txt = chart_items[row][col]
                ax.text(0.4 + 0.05 * col, 0.9 - 0.1 * row, txt, transform=ax.transAxes, fontsize=14, horizontalalignment='center', verticalalignment='center')

    plot_something(0, G1_range, G2_range, 'Range')
    plot_something(1, G1_azimuth, G2_azimuth, 'Azimuth')
    

    plt.subplots_adjust(hspace=0.5)
    #fig.set_size_inches(11.7, 8.3)
    fig.set_size_inches(15, 10.6)
    #plt.show()
    img_path = policy_path + '\\coregistration.png'
    plt.savefig(img_path)
    os.system('start ' + img_path)




if __name__ == '__main__':

    import sys
    policy_path = sys.argv[1]
    print('=== 匹配誤差報表產生器 ===')
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
        # 如果沒有 Report_3coregistration_Error 檔案的話，這筆 Policy 會被跳過

        coregistration_Error_file = policy_path + '\\' + policy_dir + '\\Report_3coregistration_Error.txt'
        if not os.path.isfile(coregistration_Error_file):
            print('找不到 Report_3coregistration_Error.txt, 跳過', policy_dir)
            continue
        
        print('開始處理', policy_dir, '...')#, end = '')

        export_chart(coregistration_Error_file, policy_path + '\\' + policy_dir)