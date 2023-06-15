import numpy as np
import pandas as pd
import matplotlib.pyplot as plt





def draw(file_name):
    # plt.rcParams["font.family"] = ["sans-serif"]
    # plt.rcParams["font.sans-serif"] = ["SimHei"]
    # plt.rcParams["axes.unicode_minus"] = False

    font = {
    'family': 'SimHei',
    'weight': 'light'
    }
    df = pd.read_csv('data/{}.csv'.format(file_name), header=0)
    data = df['风险等级'].value_counts()
    y = data.index.to_list()
    x = data.values.tolist()
    # 调整顺序
    y = [y[2],y[0],y[1]]
    x = [x[2],x[0],x[1]]


    fig = plt.figure(figsize=(8,2), dpi=150)
    ax = fig.add_subplot()
    # plt.text(3,15,'SimHei',color='red')
    # plt.tick_params(axis='y', color='red')

    ax.grid(True, linestyle='--', zorder=0)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_color('dimgrey')
    ax.spines['left'].set_color('dimgrey')
    bar_color = ['orangered', 'darkorange', 'gold']
    y_pos = np.arange(3)

    # x轴
    plt.xlim(0,35)
    # x轴刻度
    plt.xticks(np.arange(0,35,2),color='dimgrey')
    plt.yticks(y_pos,y,color='dimgrey',fontproperties=font)
    plt.tick_params(axis='x', color='dimgrey')
    plt.tick_params(axis='y', color='dimgrey')
    # ax.set_ylim(auto=False)

    # y轴
    # ax.set_yticks(y_pos, labels=y)
    # ax.set_yticks(y_pos, labels=y)

    ax.invert_yaxis()  # labels read top-to-bottom
    # x_label = [str(x[0])+'项', str(x[1])+'项', str(x[2])+'项']
    ax.barh(y_pos, x, align='center', color=bar_color, height=0.5, zorder=5)
    # ax.barh(y_pos, x, align='center', color=bar_color, zorder=10)

    # 添加字
    for a,b in enumerate(x):
        ax.text(b-1.5,a,str(b)+'项',va='center',fontsize=10, fontdict=font, color='white',zorder=10)


    # ax.set_xlabel('Performance')
    # ax.set_title('整体风险', fontdict=font)
    # ax.set_xticks(performance)
    plt.savefig('./fig/1.png', bbox_inches='tight')
    # plt.show()