# PD-Data-Processing-for-TJU-HVLab
注意：
	MATLAB R2019b及更老版本选择“Main_for_MATLAB_2019_and_older.m”
	MATLAB R2020a及更新版本选择“Main_for_MATLAB_2020_and_later.m” 
	
输入文件
	运行main.m后弹出资源管理器，选择输入文件

输出文件自动保存于main.m同目录下，命名：“*_out.xlsx”

	输出文件Sheet1为局部放电PRPD谱图，其中第一列为0°~360°相位，后续列为放电幅值

	输出文件Sheet2为局部放电特征量
	
	具体包含：最大放电幅值、放电次数、平均每秒放电次数、偏斜度S（正半周期、负半周期）、陡峭度K（正半周期、负半周期）
