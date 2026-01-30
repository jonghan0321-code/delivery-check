import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def run_program():
    # 1. 파일 선택 창 띄우기
    root = tk.Tk()
    root.withdraw() # 메인 창은 숨김
    
    file_path = filedialog.askopenfilename(
        title="처리할 PackageLegList 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return # 취소 시 종료

    try:
        # 2. 데이터 처리 로직 (기존 로직 동일)
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        allowed_zones = ['ATL', 'AUS', 'BOS', 'BWI', 'CLT', 'CVG', 'DFW', 'DTW', 'DULUTH', 'EWR', 'LAX', 'MCO', 'MIA', 'MSP', 'ORD', 'PHX', 'SEA', 'SFO', 'DEN']

        # 1차 분류 (XYZ/Zone)
        cond_xyz_keep = (df['PL D. Station'] != 'XYZ') | (df['PL D. Zone'].isin(allowed_zones))
        main_data = df[cond_xyz_keep].copy()
        xyz_raw_data = df[~cond_xyz_keep].copy()
        xyz_excluded_final = xyz_raw_data.sort_values(by=['PL No', 'Leg No'], ascending=[True, False]).drop_duplicates(subset=['PL No'], keep='first')

        # 상위 2개 Leg 추출
        main_data_sorted = main_data.sort_values(by=['PL No', 'Leg No'], ascending=[True, False])
        top1 = main_data_sorted.groupby('PL No').head(1).copy()
        top2 = main_data_sorted.groupby('PL No').nth(1).copy()

        # 최종 필터링 조건
        cond_match = ((top1['Leg P. Station'] != 'XYZ') & (top1['PL D. Station'] == top1['Leg P. Station'])) | ((top1['Leg P. Station'] == 'XYZ') & (top1['PL D. Station'] == top1['Leg P. Zone']))
        cond_status = ((top1['Leg Status'] == 'Scheduled') & (top1['Pickup ETD'].isna())) | (top1['Leg Status'].isna())

        final_candidates = top1[cond_match & cond_status].copy()
        excluded_others_final = top1[~(cond_match & cond_status)].copy()

        # 데이터 병합 및 정렬
        top2_info = top2[['PL No', 'Dest ATA']].rename(columns={'Dest ATA': 'Previous_Leg_Dest_ATA'})
        final_result = pd.merge(final_candidates[['PL No', 'PL D. Station']], top2_info, on='PL No', how='left')
        
        # 시간 순 정렬
        final_result['Previous_Leg_Dest_ATA'] = pd.to_datetime(final_result['Previous_Leg_Dest_ATA'], errors='coerce')
        final_result = final_result.sort_values(by=['PL D. Station', 'Previous_Leg_Dest_ATA'], ascending=[True, True])

        # 3. 저장 경로 설정 (선택한 파일과 같은 폴더에 저장)
        save_dir = os.path.dirname(file_path)
        output_file = os.path.join(save_dir, '메일 보내야 할 것들_결과.xlsx')

        # 4. 엑셀 저장 및 서식 적용
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            final_result.to_excel(writer, sheet_name='메일 보내야 할 것들', index=False)
            xyz_excluded_final.to_excel(writer, sheet_name='XYZ', index=False)
            excluded_others_final.to_excel(writer, sheet_name='배달예정&배달완료', index=False)

            workbook = writer.book
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
            fmt1 = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})
            fmt2 = workbook.add_format({'bg_color': '#EAF1FB', 'border': 1})

            ws = writer.sheets['메일 보내야 할 것들']
            for col_num, value in enumerate(final_result.columns.values):
                ws.write(0, col_num, value, header_fmt)

            current_fmt, prev_station = fmt1, None
            for row_num, (index, row) in enumerate(final_result.iterrows()):
                station = row['PL D. Station']
                if prev_station is not None and station != prev_station:
                    current_fmt = fmt2 if current_fmt == fmt1 else fmt1
                
                for col_num, value in enumerate(row):
                    if isinstance(value, pd.Timestamp):
                        value = value.strftime('%Y-%m-%d %H:%M')
                    ws.write(row_num + 1, col_num, value, current_fmt)
                prev_station = station
            ws.set_column(0, 2, 25)

        messagebox.showinfo("성공", f"작업이 완료되었습니다!\n파일 위치: {output_file}")

    except Exception as e:
        messagebox.showerror("오류 발생", f"에러가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    run_program()