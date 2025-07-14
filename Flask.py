from flask import Flask, request, render_template, send_file
from datetime import datetime , timedelta
import os
import pandas as pd
import copy
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    
    # Process the file (assuming Excel file for this example)
    # processed_path = process_file(filepath, filename)
    processed_path = process(filepath, filename)
    
    return send_file(processed_path, as_attachment=True)



def process(filepath, filename):
    input_file = filepath 
    df = pd.read_excel(input_file)  # Example: Read Excel file
    # df['Processed'] = df[df.columns[0]] * 2  # Example processing
    # processed_path = os.path.join(PROCESSED_FOLDER, filename)  # Save temporarily on server
    
    processed_filename = f'processed_{os.path.splitext(filename)[0]}.xlsx'
    processed_path = os.path.join(PROCESSED_FOLDER, processed_filename)
    # Contents2 = copy.deepcopy(Contents)
    # dict_data={}
    # for i in range(0, len(Contents2[0])):
    #     key = Contents2[0][i][0]
    #     value = Contents2[0][i][1]
    #     dict_data[key] = value

    # print(dict_data.items())
    df = pd.read_excel(input_file)
    start = df[df.iloc[:,0]== 'Target'].index
    end = df[df.iloc[:,0] == 'Summary for Columns'].index
    new_range = [[a,b+1] for a, b in zip(start, end)]
    df.iloc[new_range[0][0]:new_range[0][1]]
    def select_range(index):
      return df.iloc[new_range[index][0]:new_range[index][1]]
    datas = []
    for i in range(len(new_range)):
      datas.append(select_range(i))
    df[df.iloc[:,0] == 'Summary for Columns'].index
    def sum_location(df):
      df.iloc[:,0:3] = df.iloc[:,0:3].astype(str)
      last = df.shape[0]
      i = 0
      sum_location = []
      for i in range(last):
        if ('Summary' in str(df.iloc[i,0])) | ('summary' in str(df.iloc[i,0])):
          sum_location.append(i)
      i = 0
      j = 0
      for j in range(3):
        for i in range(last):
          if ('Summary' in str(df.iloc[i,j])) | (i in sum_location):
            i = i+1
          elif (df.iloc[i,j] == 'nan'):
            value = df.iloc[i-1,j]
            df.iloc[i,j] = value
            i = i+1
          else:
            i = i+1
        j = j+1
        i = 0
      df.drop(df.index[sum_location],inplace=True)
      return df
    for i in range(len(datas)):
      datas[i] = sum_location(datas[i])
    datas2 = copy.deepcopy(datas)
    des = []
    d = []
    for i in range(len(datas2)):
      var_name = datas2[i].iloc[0].dropna().values
      var_data = datas2[i].iloc[1].dropna().values
      for j in range(len(var_name)):
        name = var_name[j]
        data = var_data[j]
        d.append([name,data])
      des.append(d)
      d=[]
      datas2[i].drop(datas2[i].index[[0,1]],inplace=True)
      col_name = datas2[i].iloc[0].values
      datas2[i].columns = col_name
      datas2[i].drop(datas2[i].index[[0]],inplace=True)
      datas2[i].reset_index(drop=True,inplace=True)
    datas3 = copy.deepcopy(datas2)
    def my_date(df2_dc):
      df2_dc[['Date','Start Time']] = df2_dc[['Date','Start Time']].astype(str)
      Old_Date = df2_dc['Date'].tolist()
      Old_Time = df2_dc['Start Time'].tolist()

      merged_dates = []
      day_of_week = []
      for date_str, time_str in zip(Old_Date, Old_Time):
          hours, minutes = map(int, time_str.split(':'))  # Split "HH:MM" into integers

          # If hours >= 24, subtract 24 and add 1 day
          extra_day = hours // 24  # 1 if hours >= 24, else 0
          hours = hours % 24  # Adjust hours to be in 0-23 range

          # Convert to datetime object
          dt = datetime.strptime(date_str, '%d/%m/%Y') + timedelta(days=extra_day)
          dt = dt.replace(hour=hours, minute=minutes, second=0)  # Set time
          day = dt.weekday()
          merged_dates.append(dt)
          day_of_week.append(day)
      date_result = [dt.strftime('%d/%m/%Y %H:%M:%S') for dt in merged_dates]
      df2_dc['Merged_date'] = date_result
      df2_dc['New_Day_Of_Week'] = day_of_week
      days_mapping = {
          0: 'Mon',
          1: 'Tue',
          2: 'Wed',
          3: 'Thu',
          4: 'Fri',
          5: 'Sat',
          6: 'Sun'
      }
      df2_dc['New_Day_Of_Week'] = df2_dc['New_Day_Of_Week'].map(days_mapping)
      durations = df2_dc.Duration.tolist()
      New_Duration = []
      for duration in durations:
        New_Duration.append(duration.split(':')[2])
      df2_dc['New_Duration'] = New_Duration


    for i in range(len(datas3)):
      my_date(datas3[i])
    datas4 = copy.deepcopy(datas3)
    sheet_names = []
    i=0
    for i in range(1,len(des)+1):
      sheet_names.append('Sheet'+str(i))
      content = []
    Contents = []
    Brands = []
    Copylines = []
    # channels = df2_dc['Channel'].value_counts().index.tolist()

    for i in range(len(datas4)):
      mylist = datas4[i]['Channel'].values
      channels = list(dict.fromkeys(mylist))
      brand = des[i][1][1]
      copyline = des[i][2][1]
      Brands.append(brand)
      Copylines.append(copyline)
      for channel in channels:
          data = datas4[i][['New_Day_Of_Week','Merged_date','Break Position','Pos. in Break','New_Duration','Programme\Variables']].loc[datas4[i]['Channel'] == channel].sort_values('Merged_date').values.tolist()
          content.append([channel,data])
      Contents.append(content)
      content = []
    j=0
    i=0
    Contents2 = copy.deepcopy(Contents)
    dict_data={}
    All_Content = []
    for j in range(len(Contents2)):
      for i in range(0, len(Contents2[j])):
        key = Contents2[j][i][0]
        value = Contents2[j][i][1]
        dict_data[key] = value
      All_Content.append(dict_data)
      dict_data={}

    # Create a new Excel file and add a worksheet.
    with pd.ExcelWriter(processed_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        merge_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        merge_left = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 8,
            'bg_color' : '#D9D9D9'
        })
        merge_right = workbook.add_format({
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size' : 8,
            'bg_color' : '#D9D9D9'
        })
        default_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        merge_center = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        merge_left2 = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        merge_right2 = workbook.add_format({
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        content_left = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        content_right = workbook.add_format({
            'align': 'right',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        content_center = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        content_merge_left = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 8,
        })
        top_border_right = workbook.add_format({
            'top': 6,
            'bottom': 6,
            'align': 'right',
            'valign': 'vcenter',
            'font_size' : 8,
            'bold': True,
        })
        top_border_left = workbook.add_format({
            'top': 6,
            'bottom': 6,
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 8,
            'bold': True,
        })
        format_header = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 14,
        })
        format_subheader = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'font_size' : 7,
        })
        format_righthead = workbook.add_format({
            'align': 'right',
            'valign': 'vcenter',
            'font_size' : 7,
        })
        # workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()
        # worksheet.repeat_rows(0,3)
        worksheet.set_column("A:A", 13)
        worksheet.set_column("B:B", 7)
        worksheet.set_column("C:E", 3)
        worksheet.set_column("F:J", 8)
        worksheet.set_column("K:K", 9)
        worksheet.set_row(0, 10.5)
        worksheet.set_row(1, 10.5)
        worksheet.set_row(2, 10.5)
        # # Define multiple item sets
        # item_sets = {
        #     "Set 1 Header": [f"Item A{i+1}" for i in range(110)],
        #     "Set 2 Header": [f"Item B{i+1}" for i in range(30)],
        #     "Set 3 Header": [f"Item C{i+1}" for i in range(50)],
        # }

        # Initialize row counter
        # worksheet.merge_range("A1:H2",  'The Nielsen Company (Thailand) Limited.' , format_header)
        # worksheet.write("A3",  '34th Fls., United Center, 323 Silom Rd..Bangkok 10500 Tel. 0-2674-6000 Fax. 0-274-6000 Ext.5102' , format_subheader)
        # right_head = ['Daily Comercial Logs','Advertisment Activity','By Copyline']
        # worksheet.write_column("K1",right_head,format_righthead)
        # for i in range(1):
        data_C = 0
        data_C2 = 0

        # Write header cells without merging
        # worksheet.merge_range("A4:F4",  Brands[0] , top_border_left)
        # worksheet.merge_range("G4:K4", Copylines[0], top_border_right)
        # worksheet.merge_range("A4:F4",  Brands[i] , top_border_left)
        # worksheet.merge_range("G4:K4", Copylines[i], top_border_right)
        headers = ["", "", "Brk", "PIB", "Dur", "", "", "", "", ""]
        # my_index=5
        # datas = Contents
        # for data2 in datas:
        #   # worksheet.write('A'+str(my_index), data[0], merge_left)
        #   worksheet.merge_range('A'+str(my_index)+':F'+str(my_index), data2[0]+'  Brand : '+Brands[i], merge_left)
        #   worksheet.merge_range('G'+str(my_index)+':K'+str(my_index), Copylines[i], merge_right)
        #   my_index += 1
        #   #'A'+str(my_index)+':B'+str(my_index)
        #   worksheet.write_row('A'+str(my_index), headers, merge_format)
        #   worksheet.merge_range('A'+str(my_index)+':B'+str(my_index), "Date/Time", merge_center)
        #   # worksheet.merge_range("F3:J3", "Program", merge_left)
        #   worksheet.merge_range('F'+str(my_index)+':J'+str(my_index), "Program", merge_left2)
        #   worksheet.write('K'+str(my_index), 'Remark', merge_right2)
        #   my_index += 1
        a = 51
        my_index = 0
        row = 0
        first = True
        grand_total=0
        j = 0

        right_head = ['Daily Comercial Logs','Advertisment Activity','By Copyline']
        for dict_data in All_Content:

          for subheader, items in dict_data.items():
            # print('Subheader : ',subheader)
            first_header_written = True  # Track if header has been written
            data_remark = len(items)
            data_count = 0
            while items:
                # print(len(items))
                # If row is at 51, write subheader again
                if row % 51 == 0 or first:
                  #writer header
                  #row+4
                  # worksheet.merge_range("A1:H2",  'The Nielsen Company (Thailand) Limited.' , format_header)
                  # worksheet.write("A3",  '34th Fls., United Center, 323 Silom Rd..Bangkok 10500 Tel. 0-2674-6000 Fax. 0-274-6000 Ext.5102' , format_subheader)
                  # right_head = ['Daily Comercial Logs','Advertisment Activity','By Copyline']
                  # worksheet.write_column("K1",right_head,format_righthead)
                  #worksheet.merge_range("A4:F4",  Brands[0] , top_border_left)
                  #worksheet.merge_range("G4:K4", Copylines[0], top_border_right)
                  # worksheet.set_row(0, 10.5)
                  # worksheet.set_row(1, 10.5)
                  # worksheet.set_row(2, 10.5)
                  worksheet.set_row(row, 10.5)
                  worksheet.set_row(row+1, 10.5)
                  worksheet.write_column('K'+str(row+1),right_head,format_righthead)
                  worksheet.merge_range('A'+str(row+1)+':H'+str(row+2),  'The Nielsen Company (Thailand) Limited.' , format_header)
                  # print('A'+str(row+1)+':H'+str(row+2))
                  row+=2
                  worksheet.set_row(row, 10.5)
                  worksheet.write('A'+str(row+1),  '34th Fls., United Center, 323 Silom Rd..Bangkok 10500 Tel. 0-2674-6000 Fax. 0-274-6000 Ext.5102' , format_subheader)
                  row+=1
                  worksheet.merge_range('A'+str(row+1)+':F'+str(row+1),  Brands[j] , top_border_left)
                  worksheet.merge_range('G'+str(row+1)+':K'+str(row+1), Copylines[j], top_border_right)
                  row+=1
                  first=False
                  # print("ds")

                  #------------
                  row += 1 #mark
                  worksheet.merge_range('A'+str(row+1)+':F'+str(row+1), subheader+'  Brand : '+Brands[j],merge_left)
                  worksheet.merge_range('G'+str(row+1)+':K'+str(row+1), Copylines[j], merge_right)
                  # row += 1
                  # worksheet.merge_range('A'+str(row)+':B'+str(row), "Date/Time", merge_center)
                  # # worksheet.merge_range("F3:J3", "Program", merge_left)
                  # worksheet.merge_range('F'+str(row)+':J'+str(row), "Program", merge_left2)
                  # worksheet.write('K'+str(row), 'Remark', merge_right2)
                  # worksheet.write_row('A'+str(row), headers, merge_format)
                  row += 1
                  worksheet.write_row('A'+str(row+1), headers, merge_format)
                  worksheet.merge_range('A'+str(row+1)+':B'+str(row+1), "Date/Time", merge_center)
                  # worksheet.merge_range("F3:J3", "Program", merge_left)
                  worksheet.merge_range('F'+str(row+1)+':J'+str(row+1), "Program", merge_left2)
                  worksheet.write('K'+str(row+1), 'Remark', merge_right2)
                  row += 1
                  first_header_written = False  # Mark that header is written
                  # print('row : ',row)
                  # if data_count > 0:
                  #   a += 4
                  # data_count +=1
                  # print('a : ',a)
                  #-----------

                elif first_header_written:
                    row += 1 #mark
                    worksheet.merge_range('A'+str(row+1)+':F'+str(row+1), subheader+'  Brand : '+Brands[j],merge_left)
                    worksheet.merge_range('G'+str(row+1)+':K'+str(row+1), Copylines[j], merge_right)
                    # row += 1
                    # worksheet.merge_range('A'+str(row)+':B'+str(row), "Date/Time", merge_center)
                    # # worksheet.merge_range("F3:J3", "Program", merge_left)
                    # worksheet.merge_range('F'+str(row)+':J'+str(row), "Program", merge_left2)
                    # worksheet.write('K'+str(row), 'Remark', merge_right2)
                    # worksheet.write_row('A'+str(row), headers, merge_format)
                    row += 1
                    worksheet.write_row('A'+str(row+1), headers, merge_format)
                    worksheet.merge_range('A'+str(row+1)+':B'+str(row+1), "Date/Time", merge_center)
                    # worksheet.merge_range("F3:J3", "Program", merge_left)
                    worksheet.merge_range('F'+str(row+1)+':J'+str(row+1), "Program", merge_left2)
                    worksheet.write('K'+str(row+1), 'Remark', merge_right2)
                    row += 1
                    first_header_written = False  # Mark that header is written
                    # print('row : ',row)
                    # if data_count > 0:
                    #   a += 4
                    # data_count +=1
                    # print('a : ',a)


                # Write one item

                # worksheet.write(row, 0, items.pop(0)[5])
                # row += 1
                data = items.pop(0)
                #################
                date = data[1].split(' ')
                date_time = data[0]+' '+date[0]
                # worksheet.write_row('A'+str(my_index),[data[0]], content_left)
                worksheet.write_row('A'+str(row+1),[date_time], content_left)
                # worksheet.write_row('B'+str(my_index),[data[1]], content_left)
                worksheet.write_row('B'+str(row+1),[date[1]], content_center)
                worksheet.write_row('C'+str(row+1),[data[2]], content_center)
                worksheet.write_row('D'+str(row+1),[data[3]], content_center)
                worksheet.write_row('E'+str(row+1),[data[4]], content_center)
                worksheet.merge_range('F'+str(row+1)+':J'+str(row+1),data[5], content_merge_left)
                # worksheet.write_row('K'+str(my_index),[data[6]], content_center)
                # data_count += 1
                grand_total += 1
                row += 1
                if len(items)==0:
                  worksheet.merge_range('A'+str(row+1)+':K'+str(row+1), subheader+'  Brand : '+Brands[j]+'  Total '+str(data_remark)+' Spots',merge_left)
                  row += 1
                ###################
          # print(row)
          if row % 51 == 0:
            worksheet.set_row(row, 10.5)
            worksheet.set_row(row+1, 10.5)
            worksheet.write_column('K'+str(row+1),right_head,format_righthead)
            worksheet.merge_range('A'+str(row+1)+':H'+str(row+2),  'The Nielsen Company (Thailand) Limited.' , format_header)
            # print('A'+str(row+1)+':H'+str(row+2))
            row+=2
            worksheet.set_row(row, 10.5)
            worksheet.write('A'+str(row+1),  '34th Fls., United Center, 323 Silom Rd..Bangkok 10500 Tel. 0-2674-6000 Fax. 0-274-6000 Ext.5102' , format_subheader)
            row+=1
            worksheet.merge_range('A'+str(row+1)+':F'+str(row+1),  Brands[j] , top_border_left)
            worksheet.merge_range('G'+str(row+1)+':K'+str(row+1), Copylines[j], top_border_right)
            row+=1
            worksheet.merge_range('A'+str(row+1)+':K'+str(row+1), Brands[j]+' Grand Total '+str(grand_total)+' Spots',merge_left)
            row+=1
          else:
            worksheet.merge_range('A'+str(row+1)+':K'+str(row+1), Brands[j]+' Grand Total '+str(grand_total)+' Spots',merge_left)
            row+=1
              # my_sheet(Brands[i],Copylines[i],Contents[i])
          remainder = (row-row%a)+51-row
          row = row + remainder +1
          # print('remainder row : ',row)
          first = True
          j += 1
          grand_total=0
        worksheet.set_paper(9)
        # Close the workbook
        # workbook.close()
        print("Excel file created successfully.")

    return processed_path

def process_file(filepath, filename):
    df = pd.read_excel(filepath)  # Example: Read Excel file
    df['Processed'] = df[df.columns[0]] * 2  # Example processing
    
    processed_path = os.path.join(PROCESSED_FOLDER, filename)  # Save temporarily on server
    df.to_excel(processed_path, index=False)
    
    return processed_path

if __name__ == '__main__':
    app.run(debug=True)
    
