"""
可视化 Windows 可执行文件：
1. 使用 PySimpleGUI 构建简单界面，支持：
   - 选择输入 Excel 文件
   - 选择输出汇总文件路径
   - 编辑各科阈值参数（JSON 格式）
   - 点击“运行”按钮生成汇总
2. 依赖：pandas, openpyxl, PySimpleGUI
3. 打包：
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --windowed score_summary_gui.py
   ```
生成的 exe 即可独立运行，无需 Python 环境。
"""
import json
import pandas as pd
import PySimpleGUI as sg

# 默认参数
default_params = {
    '语文':    {'total': 150, 'excellent': 135, 'good': 120, 'pass':  90},
    '数学':    {'total': 150, 'excellent': 135, 'good': 120, 'pass':  90},
    '英语':    {'total': 150, 'excellent': 135, 'good': 120, 'pass':  90}
}

layout = [
    [sg.Text('输入文件（每科 sheet）'), sg.Input(key='-IN-'), sg.FileBrowse(file_types=(('Excel','*.xlsx'),))],
    [sg.Text('输出文件'), sg.Input(key='-OUT-'), sg.FileSaveAs(defaultextension='.xlsx', file_types=(('Excel','*.xlsx'),))],
    [sg.Text('阈值参数 (JSON)'), sg.Multiline(json.dumps(default_params, ensure_ascii=False, indent=2), size=(60,15), key='-PARAM-')],
    [sg.Button('运行'), sg.Button('退出')],
    [sg.StatusBar('', size=(60,1), key='-STATUS-')]
]

window = sg.Window('成绩汇总工具', layout)


def run_summary(input_file, output_file, params):
    index_labels = [
        '班级总分','参加考试人数','最高分','最低分','平均分',
        '合格人数','合格率','优秀人数','优秀率',
        '平均得分率','良好人数','良好率','综合率'
    ]
    with pd.ExcelFile(input_file) as xls:
        writer = pd.ExcelWriter(output_file, engine='openpyxl')
        subject_scores = {}
        for subj in xls.sheet_names:
            df = xls.parse(subj)
            scores = df.iloc[:,1].dropna()
            n = scores.count(); total = scores.sum(); mx = scores.max(); mn = scores.min(); avg = scores.mean()
            thresh = params.get(subj)
            if not thresh:
                raise KeyError(f"缺少科目 {subj} 的参数配置。")
            pass_n = scores[scores>=thresh['pass']].count()
            exc_n  = scores[scores>=thresh['excellent']].count()
            good_n = scores[scores>=thresh['good']].count()
            pass_rate  = pass_n/n; exc_rate = exc_n/n; good_rate=good_n/n; avg_rate=avg/thresh['total']
            comp_rate = avg_rate*0.2 + pass_rate*0.6 + exc_rate*0.1 + good_rate*0.1
            summary = pd.DataFrame({'指标': index_labels, subj: [
                total,n,mx,mn,avg, pass_n,pass_rate,exc_n,exc_rate, avg_rate,good_n,good_rate,comp_rate
            ]})
            summary.to_excel(writer, sheet_name=subj, index=False)
            subject_scores[subj] = scores.values
        # 综合汇总
        all_df = pd.DataFrame(subject_scores)
        all_df['总分'] = all_df.sum(axis=1)
        sums = all_df['总分']; N=sums.count(); T=sums.sum(); Mx=sums.max(); Mn=sums.min(); A=sums.mean()
        total_max = sum(v['total'] for v in params.values())
        total_pass = sum(v['pass'] for v in params.values())
        total_good = sum(v['good'] for v in params.values())
        total_exc  = sum(v['excellent'] for v in params.values())
        pass_n = sums[sums>=total_pass].count()
        exc_n  = sums[sums>=total_exc].count()
        good_n = sums[sums>=total_good].count()
        pass_rate  = pass_n/N; exc_rate=exc_n/N; good_rate=good_n/N; avg_rate=A/total_max
        comp_rate=avg_rate*0.2+pass_rate*0.6+exc_rate*0.1+good_rate*0.1
        overall=pd.DataFrame({'指标':index_labels,'综合': [
            T,N,Mx,Mn,A, pass_n,pass_rate,exc_n,exc_rate, avg_rate,good_n,good_rate,comp_rate
        ]})
        overall.to_excel(writer, sheet_name='综合', index=False)
        writer.save()

# 事件循环
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, '退出'):
        break
    if event == '运行':
        try:
            win_in = values['-IN-']
            win_out = values['-OUT-']
            params = json.loads(values['-PARAM-'])
            window['-STATUS-'].update('处理中...')
            run_summary(win_in, win_out, params)
            window['-STATUS-'].update('完成！文件已生成')
        except Exception as e:
            window['-STATUS-'].update(f'错误: {e}')

window.close()
