import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
df=pd.read_excel('F:\\测试用表格\\雷达图.xlsx')
df=df.set_index('性能评价指标')#将数据中的“性能评价指标”列设置为行索引
df=df.T#转置数据表格
df.index.name='品牌'#将转置后数据中行索引那一列的名称修改为’品牌‘
def plot_radar(data,feature):#自定义一个函数用于制作雷达图
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus']=False
    cols=['动力性','燃油经济性','制动性','操控稳定性','行驶平顺性','通过性','安全性','环保性']#指定各个品牌要显示的性能评价指标的名称
    colors=['green','blue','red','yellow']#为每个品牌设置图表中的显示颜色
    angles=np.linspace(0.1 * np.pi,2.1 * np.pi,len(cols),endpoint=False)#根据要显示的指标个数对图形进行等分
    angles=np.concatenate((angles,[angles[0]]))#连接刻度线数据
    cols = np.concatenate((cols, [cols[0]]))#闭合边界
    fig=plt.figure(figsize=(8,8))         #设置图表的窗口大小
    ax=fig.add_subplot(111,polar=True)  #设置图表在窗口中的显示位置。并设置坐标轴为极坐标体系
    for i,c in enumerate(feature):
        stats=data.loc[c]#获取品牌对应的指标数据
        stats=np.concatenate((stats,[stats[0]]))#连接品牌的指标数据
        ax.plot(angles,stats,'-',linewidth=6,c=colors[i],label='%s'%(c))#制作雷达图
        ax.fill(angles,stats,color=colors[i],alpha=0.25)#为雷达图填充颜色
    ax.legend()#为雷达图添加图例
    ax.set_yticklabels([])#隐藏坐标轴数据
    ax.set_thetagrids(angles*180/np.pi,cols,fontsize=16)#添加并设置数据标签
    plt.show()
    return fig
fig=plot_radar(df,['A品牌','B品牌','C品牌','D品牌'])#调用指定要函数制作雷达图
