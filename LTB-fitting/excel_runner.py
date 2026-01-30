import numpy as np
import xlwings as xw
from pathlib import Path
from output import find_intercept_turning_point

def run_excel_process(file_path):
    path_obj = Path(file_path)
    file_name = path_obj.name

    app = xw.App(visible=False)
    wb = xw.Book(file_path)

    try:
            app = xw.App(visible=False)
            wb = xw.Book(file_path)
            sheet_oridata = wb.sheets['Sheet1']

            #init
            sheet_names = ['result', 'calculator','data']
            existing_sheets = [sheet.name for sheet in wb.sheets]
            for target_sheet in sheet_names:
                    # 检查工作表是否存在
                    if target_sheet in existing_sheets:
                        # 存在则删除
                        wb.sheets[target_sheet].delete()
                    # 新增工作表
                    new_sheet = wb.sheets.add(target_sheet)

            sheet3 = wb.sheets['result']
            #三次方回归********************************************
            sheet2 = wb.sheets['calculator']
            sheet1 = wb.sheets['data']

            sheet1.range('A1').value = 'Test distance (cm)'
            sheet1.range('B1').value = 'X2'
            sheet1.range('C1').value = 'X3'
            sheet1.range('D1').value = 'X4'
            sheet1.range('E1').value = 'X5'
            sheet1.range('F1').value = 'X6'
            sheet1.range('G1').value = 'X7'
            sheet1.range('H1').value = 'X8'
            sheet1.range('I1').value = 'X9'
            sheet1.range('J1').value = 'Measured E-field (V/m)'
            sheet1.range('A2:A17').value = '=Sheet1!A2:A17'
            sheet1.range('J2:J17').value = '=Sheet1!B2:B17'
            for i in range(2,10):
                for j in range(2,18):
                    cell = sheet1.range((j, i))
                    cell.formula = f'=POWER(A{j}, {i})'
            
            
            
            output_range = sheet2.range('A1:D5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:C17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_3=sheet2.range('B4').value
            Intercept_3=sheet2.range('D1').value
            Stander_Error_Intercept_3=sheet2.range('D2').value

            #标准误差
            Stander_Error_3=sheet2.range('B3').value
            #X Variable 1
            X1_3=sheet2.range('C1').value
            Stander_X1_3=sheet2.range('C2').value
            #X Variable 2
            X2_3=sheet2.range('B1').value
            Stander_X2_3=sheet2.range('B2').value
            #X Variable 3
            X3_3=sheet2.range('A1').value
            Stander_X3_3=sheet2.range('A2').value
            # t 临界值​​
            sheet2.range('A7').formula_array = f'=T.INV.2T(0.05, {Cancha_3})'
            t_3=sheet2.range('A7').value
            #lower95%
            lower_3=Intercept_3-(Stander_Error_Intercept_3*t_3)
            sheet2.range('B6').formula = 'lower 95%'
            sheet2.range('B7').formula = lower_3
            #upper95%
            upper_3=Intercept_3+(Stander_Error_Intercept_3*t_3)
            sheet2.range('C6').formula = 'upper 95%'
            sheet2.range('C7').formula = upper_3

            #R Square
            R_3=sheet2.range('A3').value
            #Multiple R
            MR_3=np.sqrt(R_3)
            sheet2.range('A8').formula = 'Multiple R'
            sheet2.range('A9').formula = MR_3
            #Adjusted R Square
            AR_3=1-((1-R_3)*(16-1)/(16-3-1))
            sheet2.range('B8').formula = 'Adjusted R Square'
            sheet2.range('B9').formula = AR_3





            #*********************************************

            #四次方回归
            output_range = sheet2.range('F1:J5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:D17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_4=sheet2.range('G4').value
            Intercept_4=sheet2.range('J1').value
            Stander_Error_Intercept_4=sheet2.range('J2').value

            #标准误差
            Stander_Error_4=sheet2.range('G3').value
            #X Variable 1
            X1_4=sheet2.range('I1').value
            Stander_X1_4=sheet2.range('I2').value
            #X Variable 2
            X2_4=sheet2.range('H1').value
            Stander_X2_4=sheet2.range('H2').value
            #X Variable 3
            X3_4=sheet2.range('G1').value
            Stander_X3_4=sheet2.range('G2').value
            #X Variable 4
            X4_4=sheet2.range('F1').value
            Stander_X4_4=sheet2.range('F2').value

            # t 临界值​​
            sheet2.range('F7').formula_array = f'=T.INV.2T(0.05, {Cancha_4})'
            t_4=sheet2.range('F7').value
            #lower95%
            lower_4=Intercept_4-(Stander_Error_Intercept_4*t_4)
            sheet2.range('G6').formula = 'lower 95%'
            sheet2.range('G7').formula = lower_4
            #upper95%
            upper_4=Intercept_4+(Stander_Error_Intercept_4*t_4)
            sheet2.range('H6').formula = 'upper 95%'
            sheet2.range('H7').formula = upper_4

            #R Square
            R_4=sheet2.range('F3').value
            #Multiple R
            MR_4=np.sqrt(R_4)
            sheet2.range('F8').formula = 'Multiple R'
            sheet2.range('F9').formula = MR_4
            #Adjusted R Square
            AR_4=1-((1-R_4)*(16-1)/(16-4-1))
            sheet2.range('G8').formula = 'Adjusted R Square'
            sheet2.range('G9').formula = AR_4





            #五次方回归：
            output_range = sheet2.range('L1:Q5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:E17,TRUE,TRUE)'


            wb.app.calculate()

            Cancha_5=sheet2.range('M4').value
            Intercept_5=sheet2.range('Q1').value
            Stander_Error_Intercept_5=sheet2.range('Q2').value

            #标准误差
            Stander_Error_5=sheet2.range('M3').value
            #X Variable 1
            X1_5=sheet2.range('P1').value
            Stander_X1_5=sheet2.range('P2').value
            #X Variable 2
            X2_5=sheet2.range('O1').value
            Stander_X2_5=sheet2.range('O2').value
            #X Variable 3
            X3_5=sheet2.range('N1').value
            Stander_X3_5=sheet2.range('N2').value
            #X Variable 4
            X4_5=sheet2.range('M1').value
            Stander_X4_5=sheet2.range('M2').value
            #X Variable 5
            X5_5=sheet2.range('L1').value
            Stander_X5_5=sheet2.range('L2').value

            # t 临界值​​
            sheet2.range('L7').formula_array = f'=T.INV.2T(0.05, {Cancha_5})'
            t_5=sheet2.range('L7').value

            #lower95%
            lower_5=Intercept_5-(Stander_Error_Intercept_5*t_5)
            sheet2.range('M6').formula = 'lower 95%'
            sheet2.range('M7').formula = lower_5
            #upper95%
            upper_5=Intercept_5+(Stander_Error_Intercept_5*t_5)
            sheet2.range('N6').formula = 'upper 95%'
            sheet2.range('N7').formula = upper_5

            #R Square
            R_5=sheet2.range('L3').value
            #Multiple R
            MR_5=np.sqrt(R_5)
            sheet2.range('L8').formula = 'Multiple R'
            sheet2.range('L9').formula = MR_5
            #Adjusted R Square
            AR_5=1-((1-R_5)*(16-1)/(16-5-1))
            sheet2.range('M8').formula = 'Adjusted R Square'
            sheet2.range('M9').formula = AR_5




            # 六次方回归
            output_range = sheet2.range('S1:Y5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:F17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_6 = sheet2.range('T4').value
            Intercept_6 = sheet2.range('Y1').value
            Stander_Error_Intercept_6 = sheet2.range('Y2').value

            # 标准误差
            Stander_Error_6 = sheet2.range('T3').value
            # X Variable 1
            X1_6 = sheet2.range('X1').value
            Stander_X1_6 = sheet2.range('X2').value
            # X Variable 2
            X2_6 = sheet2.range('W1').value
            Stander_X2_6 = sheet2.range('W2').value
            # X Variable 3
            X3_6 = sheet2.range('V1').value
            Stander_X3_6 = sheet2.range('V2').value
            # X Variable 4
            X4_6 = sheet2.range('U1').value
            Stander_X4_6 = sheet2.range('U2').value
            # X Variable 5
            X5_6 = sheet2.range('T1').value
            Stander_X5_6 = sheet2.range('T2').value
            # X Variable 6
            X6_6 = sheet2.range('S1').value
            Stander_X6_6 = sheet2.range('S2').value

            # t 临界值
            sheet2.range('S7').formula_array = f'=T.INV.2T(0.05, {Cancha_6})'
            t_6 = sheet2.range('S7').value

            # lower95%
            lower_6 = Intercept_6 - (Stander_Error_Intercept_6 * t_6)
            sheet2.range('T6').formula = 'lower 95%'
            sheet2.range('T7').formula = lower_6
            # upper95%
            upper_6 = Intercept_6 + (Stander_Error_Intercept_6 * t_6)
            sheet2.range('U6').formula = 'upper 95%'
            sheet2.range('U7').formula = upper_6

            # R Square
            R_6 = sheet2.range('S3').value
            # Multiple R
            MR_6 = np.sqrt(R_6)
            sheet2.range('S8').formula = 'Multiple R'
            sheet2.range('S9').formula = MR_6
            # Adjusted R Square
            AR_6 = 1 - ((1 - R_6) * (16 - 1) / (16 - 6 - 1))
            sheet2.range('T8').formula = 'Adjusted R Square'
            sheet2.range('T9').formula = AR_6



            # 七次方回归 (需要8列：AA到AH)
            output_range = sheet2.range('AA1:AH5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:G17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_7 = sheet2.range('AB4').value
            Intercept_7 = sheet2.range('AH1').value
            Stander_Error_Intercept_7 = sheet2.range('AH2').value

            # 标准误差
            Stander_Error_7 = sheet2.range('AB3').value
            # X Variable 1
            X1_7 = sheet2.range('AG1').value
            Stander_X1_7 = sheet2.range('AG2').value
            # X Variable 2
            X2_7 = sheet2.range('AF1').value
            Stander_X2_7 = sheet2.range('AF2').value
            # X Variable 3
            X3_7 = sheet2.range('AE1').value
            Stander_X3_7 = sheet2.range('AE2').value
            # X Variable 4
            X4_7 = sheet2.range('AD1').value
            Stander_X4_7 = sheet2.range('AD2').value
            # X Variable 5
            X5_7 = sheet2.range('AC1').value
            Stander_X5_7 = sheet2.range('AC2').value
            # X Variable 6
            X6_7 = sheet2.range('AB1').value
            Stander_X6_7 = sheet2.range('AB2').value
            # X Variable 7
            X7_7 = sheet2.range('AA1').value
            Stander_X7_7 = sheet2.range('AA2').value

            # t 临界值
            sheet2.range('AA7').formula_array = f'=T.INV.2T(0.05, {Cancha_7})'
            t_7 = sheet2.range('AA7').value

            # lower95%
            lower_7 = Intercept_7 - (Stander_Error_Intercept_7 * t_7)
            sheet2.range('AB6').formula = 'lower 95%'
            sheet2.range('AB7').formula = lower_7
            # upper95%
            upper_7 = Intercept_7 + (Stander_Error_Intercept_7 * t_7)
            sheet2.range('AC6').formula = 'upper 95%'
            sheet2.range('AC7').formula = upper_7

            # R Square
            R_7 = sheet2.range('AA3').value
            # Multiple R
            MR_7 = np.sqrt(R_7)
            sheet2.range('AA8').formula = 'Multiple R'
            sheet2.range('AA9').formula = MR_7
            # Adjusted R Square
            AR_7 = 1 - ((1 - R_7) * (16 - 1) / (16 - 7 - 1))
            sheet2.range('AB8').formula = 'Adjusted R Square'
            sheet2.range('AB9').formula = AR_7


            #八次方回归
            output_range = sheet2.range('AJ1:AR5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:H17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_8 = sheet2.range('AK4').value  # 自由度在第4行第2列
            Intercept_8 = sheet2.range('AR1').value  # 截距在第1行最后一列
            Stander_Error_Intercept_8 = sheet2.range('AR2').value  # 截距标准误差在第2行最后一列

            # 标准误差 (在第3行第2列)
            Stander_Error_8 = sheet2.range('AK3').value
            # X Variable 1 (第1行倒数第2列)
            X1_8 = sheet2.range('AQ1').value
            Stander_X1_8 = sheet2.range('AQ2').value
            # X Variable 2 (第1行倒数第3列)
            X2_8 = sheet2.range('AP1').value
            Stander_X2_8 = sheet2.range('AP2').value
            # X Variable 3 (第1行倒数第4列)
            X3_8 = sheet2.range('AO1').value
            Stander_X3_8 = sheet2.range('AO2').value
            # X Variable 4 (第1行倒数第5列)
            X4_8 = sheet2.range('AN1').value
            Stander_X4_8 = sheet2.range('AN2').value
            # X Variable 5 (第1行倒数第6列)
            X5_8 = sheet2.range('AM1').value
            Stander_X5_8 = sheet2.range('AM2').value
            # X Variable 6 (第1行倒数第7列)
            X6_8 = sheet2.range('AL1').value
            Stander_X6_8 = sheet2.range('AL2').value
            # X Variable 7 (第1行倒数第8列)
            X7_8 = sheet2.range('AK1').value
            Stander_X7_8 = sheet2.range('AK2').value
            # X Variable 8 (第1行第1列)
            X8_8 = sheet2.range('AJ1').value
            Stander_X8_8 = sheet2.range('AJ2').value

            # t 临界值
            sheet2.range('AJ7').formula_array = f'=T.INV.2T(0.05, {Cancha_8})'
            t_8 = sheet2.range('AJ7').value

            # lower95%
            lower_8 = Intercept_8 - (Stander_Error_Intercept_8 * t_8)
            sheet2.range('AK6').formula = 'lower 95%'
            sheet2.range('AK7').formula = lower_8
            # upper95%
            upper_8 = Intercept_8 + (Stander_Error_Intercept_8 * t_8)
            sheet2.range('AL6').formula = 'upper 95%'
            sheet2.range('AL7').formula = upper_8

            # R Square (在第3行第1列)
            R_8 = sheet2.range('AJ3').value
            # Multiple R
            MR_8 = np.sqrt(R_8)
            sheet2.range('AJ8').formula = 'Multiple R'
            sheet2.range('AJ9').formula = MR_8
            # Adjusted R Square
            AR_8 = 1 - ((1 - R_8) * (16 - 1) / (16 - 8 - 1))
            sheet2.range('AK8').formula = 'Adjusted R Square'
            sheet2.range('AK9').formula = AR_8


            # 九次方回归 (需要10列：AT到BC)
            output_range = sheet2.range('AT1:BC5')
            output_range.formula_array = f'=LINEST({sheet1.name}!J2:J17, {sheet1.name}!A2:I17,TRUE,TRUE)'

            wb.app.calculate()

            Cancha_9 = sheet2.range('AU4').value  # 自由度在第4行第2列
            Intercept_9 = sheet2.range('BC1').value  # 截距在第1行最后一列
            Stander_Error_Intercept_9 = sheet2.range('BC2').value  # 截距标准误差在第2行最后一列

            # 标准误差 (在第3行第2列)
            Stander_Error_9 = sheet2.range('AU3').value
            # X Variable 1 (第1行倒数第2列)
            X1_9 = sheet2.range('BB1').value
            Stander_X1_9 = sheet2.range('BB2').value
            # X Variable 2 (第1行倒数第3列)
            X2_9 = sheet2.range('BA1').value
            Stander_X2_9 = sheet2.range('BA2').value
            # X Variable 3 (第1行倒数第4列)
            X3_9 = sheet2.range('AZ1').value
            Stander_X3_9 = sheet2.range('AZ2').value
            # X Variable 4 (第1行倒数第5列)
            X4_9 = sheet2.range('AY1').value
            Stander_X4_9 = sheet2.range('AY2').value
            # X Variable 5 (第1行倒数第6列)
            X5_9 = sheet2.range('AX1').value
            Stander_X5_9 = sheet2.range('AX2').value
            # X Variable 6 (第极1行倒数第7列)
            X6_9 = sheet2.range('AW1').value
            Stander_X6_9 = sheet2.range('AW2').value
            # X Variable 7 (第1行倒数第8列)
            X7_9 = sheet2.range('AV1').value
            Stander_X7_9 = sheet2.range('AV2').value
            # X Variable 8 (第1行倒数第9列)
            X8_9 = sheet2.range('AU1').value
            Stander_X8_9 = sheet2.range('AU2').value
            # X Variable 9 (第1行第1列)
            X9_9 = sheet2.range('AT1').value
            Stander_X9_9 = sheet2.range('AT2').value

            # t 临界值
            sheet2.range('AT7').formula_array = f'=T.INV.2T(0.05, {Cancha_9})'
            t_9 = sheet2.range('AT7').value

            # lower95%
            lower_9 = Intercept_9 - (Stander_Error_Intercept_9 * t_9)
            sheet2.range('AU6').formula = 'lower 95%'
            sheet2.range('AU7').formula = lower_9
            # upper95%
            upper_9 = Intercept_9 + (Stander_Error_Intercept_9 * t_9)
            sheet2.range('AV6').formula = 'upper 95%'
            sheet2.range('AV7').formula = upper_9

            # R Square (在第3行第1列)
            R_9 = sheet2.range('AT3').value
            # Multiple R
            MR_9 = np.sqrt(R_9)
            sheet2.range('AT8').formula = 'Multiple R'
            sheet2.range('AT9').formula = MR_9
            # Adjusted R Square
            AR_9 = 1 - ((1 - R_9) * (16 - 1) / (16 - 9 - 1))
            sheet2.range('AU8').formula = 'Adjusted R Square'
            sheet2.range('AU9').formula = AR_9


            intercepts_list=[Intercept_3, Intercept_4, Intercept_5, Intercept_6, Intercept_7, Intercept_8, Intercept_9]
            Multiple_R_list=[MR_3, MR_4, MR_5, MR_6, MR_7, MR_8, MR_9]
            R_Square_list=[R_3, R_4, R_5, R_6, R_7, R_8, R_9]
            Adjusted_R_Square_list=[AR_3, AR_4, AR_5, AR_6, AR_7, AR_8, AR_9]
            Stander_Error_list=[Stander_Error_3, Stander_Error_4, Stander_Error_5, Stander_Error_6, Stander_Error_7, Stander_Error_8, Stander_Error_9]  
            #观测值=16
            upper_list=[upper_3, upper_4, upper_5, upper_6, upper_7, upper_8, upper_9]
            lower_list=[lower_3, lower_4, lower_5, lower_6, lower_7, lower_8, lower_9]

            X1_Variable_list=[X1_3,X1_4,X1_5,X1_6,X1_7,X1_8,X1_9]
            X2_Variable_list=[X2_3,X2_4,X2_5,X2_6,X2_7,X2_8,X2_9]
            X3_Variable_list=[X3_3,X3_4,X3_5,X3_6,X3_7,X3_8,X3_9]
            X4_Variable_list=[-1,X4_4,X4_5,X4_6,X4_7,X4_8,X4_9]
            X5_Variable_list=[-1,-1,X5_5,X5_6,X5_7,X5_8,X5_9]
            X6_Variable_list=[-1,-1,-1,X6_6,X6_7,X6_8,X6_9]
            X7_Variable_list=[-1,-1,-1,-1,X7_7,X7_8,X7_9]   
            X8_Variable_list=[-1,-1,-1,-1,-1,X8_8,X8_9]
            X9_Variable_list=[-1,-1,-1,-1,-1,-1,X9_9]



            Stander_Error_count, stop_index, n_value = find_intercept_turning_point(Stander_Error_list)

            print(f"\n总循环比较次数: {Stander_Error_count}")
            print(f"停止时在列表中的索引: {stop_index}")
            print(f"停止时对应的 n 值 (Intercept_n): {n_value}")

            row=1
            row_X=2
            for i in range(0,7):
                sheet3.range(f'A{row}').value = f'{i+3}次方回归'
                row+=1
                sheet3.range(f'A{row}').value = 'Multiple_R'
                sheet3.range(f'B{row}').value = Multiple_R_list[i]
                row+=1
                sheet3.range(f'A{row}').value = 'R_Square'
                sheet3.range(f'B{row}').value = R_Square_list[i]
                row+=1
                sheet3.range(f'A{row}').value = 'Adjusted_R_Square'
                sheet3.range(f'B{row}').value = Adjusted_R_Square_list[i]
                row+=1
                sheet3.range(f'A{row}').value = '标准误差'
                sheet3.range(f'B{row}').value = Stander_Error_list[i]
                if (i+3==n_value):
                    sheet3.range(f'B{row}').color = (144, 238, 144)
                row+=1
                sheet3.range(f'A{row}').value = '观测值'
                sheet3.range(f'B{row}').value = '16'
                row+=2
                sheet3.range(f'A{row}').value = 'Upper 95%'
                sheet3.range(f'B{row}').value = upper_list[i]
                row+=1
                sheet3.range(f'A{row}').value = 'Lower 95%'
                sheet3.range(f'B{row}').value = lower_list[i]
                row+=8
                sheet3.range(f'D{row_X}').value = 'Intercept'
                sheet3.range(f'D{row_X}').color = (255, 255, 0)
                sheet3.range(f'E{row_X}').value = intercepts_list[i]
                sheet3.range(f'E{row_X}').color = (255, 255, 0)
                if (i+3==n_value):
                    if(intercepts_list[i]>83):
                        sheet3.range(f'F{row_X}').color = (255, 0, 0)
                row_X+=1
                
                for j in range(1,i+4):
                    sheet3.range(f'D{row_X}').value = f'X{j} Variable'
                    if j==1:
                        sheet3.range(f'E{row_X}').value = X1_Variable_list[i]
                    elif j==2:
                        sheet3.range(f'E{row_X}').value = X2_Variable_list[i]
                    elif j==3:
                        sheet3.range(f'E{row_X}').value = X3_Variable_list[i]
                    elif j==4:
                        sheet3.range(f'E{row_X}').value = X4_Variable_list[i]
                    elif j==5:
                        sheet3.range(f'E{row_X}').value = X5_Variable_list[i]
                    elif j==6:
                        sheet3.range(f'E{row_X}').value = X6_Variable_list[i]
                    elif j==7:
                        sheet3.range(f'E{row_X}').value = X7_Variable_list[i]
                    elif j==8:
                        sheet3.range(f'E{row_X}').value = X8_Variable_list[i]
                    elif j==9:
                        sheet3.range(f'E{row_X}').value = X9_Variable_list[i]

                    row_X+=1

                row_X=row+1

            sheet2_row=12    
            num_columns_to_offset =  4
            num_rows_to_fill = 17 # 你希望溢出的行数，例如Sheet1中A1:A17是17行

            X_3Variables_list = [X1_3, X2_3, X3_3]
            X_4Variables_list = [X1_4, X2_4, X3_4, X4_4]
            X_5Variables_list = [X1_5, X2_5, X3_5, X4_5, X5_5]
            X_6Variables_list = [X1_6, X2_6, X3_6, X4_6, X5_6, X6_6]
            X_7Variables_list = [X1_7, X2_7, X3_7, X4_7, X5_7, X6_7, X7_7]
            X_8Variables_list = [X1_8, X2_8, X3_8, X4_8, X5_8, X6_8, X7_8, X8_8]
            X_9Variables_list = [X1_9, X2_9, X3_9, X4_9, X5_9, X6_9, X7_9, X8_9, X9_9]

            for i in range(0, 7):
                # 定义目标区域的起始单元格和结束单元格
                start_cell_testdistan = (sheet2_row, 1 + i * num_columns_to_offset)
                end_cell_testdistan = (sheet2_row + num_rows_to_fill - 1, 1 + i * num_columns_to_offset) # 减去1因为起始行已经包含
                testdistan_range = sheet2.range(start_cell_testdistan, end_cell_testdistan) # 创建一个区域对象

                start_cell_measured_e = (sheet2_row, 2 + i * num_columns_to_offset)
                end_cell_measured_e = (sheet2_row + num_rows_to_fill - 1, 2 + i * num_columns_to_offset)
                measured_e_range = sheet2.range(start_cell_measured_e, end_cell_measured_e)

                start_cell_Cubic = (sheet2_row, 3 + i * num_columns_to_offset)
                end_cell_Cubic = (sheet2_row, 3 + i * num_columns_to_offset)
                Cubic_range = sheet2.range(start_cell_Cubic, end_cell_Cubic)
                
                
                
            #写入回归y值
                # 将公式数组赋值给这个多单元格区域
                testdistan_range.formula_array = f'={sheet1.name}!A1:A17' # xlwings可能会自动处理跨工作表的引用
                measured_e_range.formula_array = f'={sheet1.name}!J1:J17'
                Cubic_range.formula_array = 'Order(E-field (V/m))'
                Td='RC[-2]'

                if i==0:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_3}+{X_3Variables_list[0]}*{Td}+{X_3Variables_list[1]}*POWER({Td},2)+{X_3Variables_list[2]}*POWER({Td},3)'
                elif i==1:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_4}+{X_4Variables_list[0]}*{Td}+{X_4Variables_list[1]}*POWER({Td},2)+{X_4Variables_list[2]}*POWER({Td},3)+{X_4Variables_list[3]}*POWER({Td},4)'
                elif i==2:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_5}+{X_5Variables_list[0]}*{Td}+{X_5Variables_list[1]}*POWER({Td},2)+{X_5Variables_list[2]}*POWER({Td},3)+{X_5Variables_list[3]}*POWER({Td},4)+{X_5Variables_list[4]}*POWER({Td},5)'
                elif i==3:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_6}+{X_6Variables_list[0]}*{Td}+{X_6Variables_list[1]}*POWER({Td},2)+{X_6Variables_list[2]}*POWER({Td},3)+{X_6Variables_list[3]}*POWER({Td},4)+{X_6Variables_list[4]}*POWER({Td},5)+{X_6Variables_list[5]}*POWER({Td},6)'
                elif i==4:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_7}+{X_7Variables_list[0]}*{Td}+{X_7Variables_list[1]}*POWER({Td},2)+{X_7Variables_list[2]}*POWER({Td},3)+{X_7Variables_list[3]}*POWER({Td},4)+{X_7Variables_list[4]}*POWER({Td},5)+{X_7Variables_list[5]}*POWER({Td},6)+{X_7Variables_list[6]}*POWER({Td},7)'
                elif i==5:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_8}+{X_8Variables_list[0]}*{Td}+{X_8Variables_list[1]}*POWER({Td},2)+{X_8Variables_list[2]}*POWER({Td},3)+{X_8Variables_list[3]}*POWER({Td},4)+{X_8Variables_list[4]}*POWER({Td},5)+{X_8Variables_list[5]}*POWER({Td},6)+{X_8Variables_list[6]}*POWER({Td},7)+{X_8Variables_list[7]}*POWER({Td},8)'
                elif i==6:
                    for order_temp_row in range(0,num_rows_to_fill-1):
                        order_range = sheet2.range(sheet2_row+1+order_temp_row, 3 + i * num_columns_to_offset)
                        order_range.formula = f'={Intercept_9}+{X_9Variables_list[0]}*{Td}+{X_9Variables_list[1]}*POWER({Td},2)+{X_9Variables_list[2]}*POWER({Td},3)+{X_9Variables_list[3]}*POWER({Td},4)+{X_9Variables_list[4]}*POWER({Td},5)+{X_9Variables_list[5]}*POWER({Td},6)+{X_9Variables_list[6]}*POWER({Td},7)+{X_9Variables_list[7]}*POWER({Td},8)+{X_9Variables_list[8]}*POWER({Td},9)'
                # i从0开始，所以需要加3来表示次数

            xy1 = f"y={Intercept_3}{'+' if X1_3 >= 0 else ''}{X1_3}*x{'+' if X2_3 >= 0 else ''}{X2_3}*x^2{'+' if X3_3 >= 0 else ''}{X3_3}*x^3"
            xy2 = f"y={Intercept_4}{'+' if X1_4 >= 0 else ''}{X1_4}*x{'+' if X2_4 >= 0 else ''}{X2_4}*x^2{'+' if X3_4 >= 0 else ''}{X3_4}*x^3{'+' if X4_4 >= 0 else ''}{X4_4}*x^4"
            xy3 = f"y={Intercept_5}{'+' if X1_5 >= 0 else ''}{X1_5}*x{'+' if X2_5 >= 0 else ''}{X2_5}*x^2{'+' if X3_5 >= 0 else ''}{X3_5}*x^3{'+' if X4_5 >= 0 else ''}{X4_5}*x^4{'+' if X5_5 >= 0 else ''}{X5_5}*x^5"
            xy4 = f"y={Intercept_6}{'+' if X1_6 >= 0 else ''}{X1_6}*x{'+' if X2_6 >= 0 else ''}{X2_6}*x^2{'+' if X3_6 >= 0 else ''}{X3_6}*x^3{'+' if X4_6 >= 0 else ''}{X4_6}*x^4{'+' if X5_6 >= 0 else ''}{X5_6}*x^5{'+' if X6_6 >= 0 else ''}{X6_6}*x^6"
            xy5 = f"y={Intercept_7}{'+' if X1_7 >= 0 else ''}{X1_7}*x{'+' if X2_7 >= 0 else ''}{X2_7}*x^2{'+' if X3_7 >= 0 else ''}{X3_7}*x^3{'+' if X4_7 >= 0 else ''}{X4_7}*x^4{'+' if X5_7 >= 0 else ''}{X5_7}*x^5{'+' if X6_7 >= 0 else ''}{X6_7}*x^6{'+' if X7_7 >= 0 else ''}{X7_7}*x^7"
            xy6 = f"y={Intercept_8}{'+' if X1_8 >= 0 else ''}{X1_8}*x{'+' if X2_8 >= 0 else ''}{X2_8}*x^2{'+' if X3_8 >= 0 else ''}{X3_8}*x^3{'+' if X4_8 >= 0 else ''}{X4_8}*x^4{'+' if X5_8 >= 0 else ''}{X5_8}*x^5{'+' if X6_8 >= 0 else ''}{X6_8}*x^6{'+' if X7_8 >= 0 else ''}{X7_8}*x^7{'+' if X8_8 >= 0 else ''}{X8_8}*x^8"
            xy7 = f"y={Intercept_9}{'+' if X1_9 >= 0 else ''}{X1_9}*x{'+' if X2_9 >= 0 else ''}{X2_9}*x^2{'+' if X3_9 >= 0 else ''}{X3_9}*x^3{'+' if X4_9 >= 0 else ''}{X4_9}*x^4{'+' if X5_9 >= 0 else ''}{X5_9}*x^5{'+' if X6_9 >= 0 else ''}{X6_9}*x^6{'+' if X7_9 >= 0 else ''}{X7_9}*x^7{'+' if X8_9 >= 0 else ''}{X8_9}*x^8{'+' if X9_9 >= 0 else ''}{X9_9}*x^9"
            xy_list=[xy1,xy2,xy3,xy4,xy5,xy6,xy7]

            #画表in sheet3:result
            for n in range(0, 7):
                chart = sheet3.charts.add()

                chart.api[0].Width = 500  # 将图表宽度设置为600点
                chart.chart_type = 'xy_scatter_smooth'  # 设置图表类型

                start_cell_chart = (sheet2_row, 1 + n * num_columns_to_offset)
                end_cell_chart = (sheet2_row + num_rows_to_fill - 1, 3 + n * num_columns_to_offset)
                
                chart.set_source_data(sheet2.range(start_cell_chart, end_cell_chart)) 
                chart.api[1].SetElement(2) 
                chart.api[1].ChartTitle.Text = xy_list[n]  # 设置图表标题
                chart_title_font = chart.api[1].ChartTitle.Font
                chart_title_font.Name = "微软雅黑"  # 设置字体
                chart_title_font.Size = 8        # 设置字号


                chart.api[1].Axes(1).HasTitle = True  # 设置X轴显示标题 (1代表横坐标轴)
                chart.api[1].Axes(1).AxisTitle.Text = "Test distance (cm)"
                chart.api[1].Axes(1).MajorUnit = 5

                chart.api[1].Axes(2).HasTitle = True  # 设置Y轴显示标题 (2代表纵坐标轴)
                chart.api[1].Axes(2).AxisTitle.Text = "E-field (V/m)"

                chart.top = sheet3.range(f'G{n*16+1}').top
                chart.left = sheet3.range(f'G{n*16+2}').left

                # 强制刷新Excel界面以显示更新的数据
            wb.app.calculate()
            wb.save()
            return file_name

    finally:
        try:
            wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass
