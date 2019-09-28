class Main:
    def __init__(self):
        # self.Import_Required_Library()
        self.url  =  "https://docs.google.com/spreadsheets/d/1HYjfEe3aCbufbqIXKs0Xz-gfoQNztGhCN1ivx0gZXnc/export?format = xlsx"    
        self.Clone_the_dataset_to_this_machine(self.url)
        self.data  =  self.creat_data("Student Gradebook.xlsx")
        self.data  =  self.Add_Month_column(self.data)
        #self.today  =  date.today()
        given_name = "Vishal"
        given_month = "August"
        cwd  =  os.getcwd()
        try:
            os.mkdir(cwd+"\\" + given_name+"_"+given_month)
        except:
            None
        self.file_loc  =  cwd+"\\" + given_name+"_"+given_month + "\\"
        self.Start_making_pdf_of(given_name, given_month)
        
        
    def Import_Required_Library(self):
        from datetime import date
        from datetime import datetime
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        from reportlab.lib.pagesizes import A4
        import pandas as pd
        import numpy as np
        import openpyxl
        import requests
        import seaborn as sns
        import matplotlib.pyplot as plt
        from math import pi
        import os
        
        
    def Clone_the_dataset_to_this_machine(self, url):
        a  =  requests.get(url)
        resp  =  requests.get(url)
        output  =  open('Student Gradebook.xlsx',  'wb')
        output.write(resp.content)
        output.close()
        
        
    def creat_data(self, file):
        wb  =  openpyxl.load_workbook(file) 
        Number_of_sheets  =  len(wb.sheetnames)
        for i in range(len(wb.sheetnames)):
            try:
                data  =  pd.concat([data,  pd.read_excel(file, sheet_name = i)])
            except:
                data  =  pd.read_excel(file, sheet_name = 0)
            #print(data.shape)
        return data

        
    def Add_Month_column(self, data):
        date_list  =  []
        day_list  =  []
        year_list  =  []
        a = 0
        for i in data["Date"]:
            try:
                date_list.append(i.strftime("%B"))
                day_list.append(i.strftime("%d"))
                year_list.append(i.strftime("%Y"))
                a  +=  1
            except:
                try:
                    date_list.append(datetime.strptime(str(i), "%d/%m/%Y").strftime("%B"))
                    day_list.append(datetime.strptime(str(i), "%d/%m/%Y").strftime("%d"))
                    year_list.append(datetime.strptime(str(i), "%d/%m/%Y").strftime("%Y"))
                    a +=  1
                except:
                    print("Interrupt in a line number : - ", a, "which is having a entry of:-", i)
                    break
        data["Date(only)"]  =  day_list
        data["Month"]  =  date_list
        data["Year"]  =  year_list
        return data
    def Return_prepared_data(self):
        print(self.data.shape)
        
    def Start_making_pdf_of(self, name, month):
        self.Name  =  name
        self.College  =  "GCELT"
        self.Year  =  "First"
        self.Month  =  month
        self.Date  =  str(date.today().strftime("%d/%m/%Y"))
        self.Email  =  "sakibmondal7@gmail.com"
        self.Number_of_task_wins  =  self.number_of_task_wins(self.Name, self.Month)
        self.Rank_among_the_class  =  self.rank_of_the_student(self.Name, self.Month)
        self.Late_submition_ratio  =  self.late_Submition_Ratio(self.Name, self.Month)
        self.Percentage  =  self.percentage_of_the_student(self.Name, self.Month)
        self.Percentile   =  self.percentile_of_the_student(self.Name, self.Month)
        self.Table_Content  =  self.table_Content(self.Name, self.Month)
        self.Table_summary  =  self.table_summary(self.Name, self.Month)
        self.main_of_pdf(self.Name, self.Month)
        
        
    def number_of_task_wins(self, name, month):
        working_data  =  self.data
        working_data  =  working_data[working_data["Student"]  ==  name]
        working_data  =  working_data[working_data["Month"]  ==  month]
        working_data  =  working_data[working_data["Task Winner"]  ==  1]
        return str(working_data['id'].count())
    
    def rank_of_the_student(self, name, month):
        working_data  =  self.data
        return str(5)
    
    def late_Submition_Ratio(self, name, month):
        working_data  =  self.data
        return str(0.5)
        
    def percentage_of_the_student(self, name, month):
        working_data  =  self.data
        return str(88)
    
    def percentile_of_the_student(self, name, month):
        working_data  =  self.data
        return str(95)
    
    def table_Content(self, name, month):
        working_data  =  self.data
        a  =  [["1Data Analytics", 30.0, 26.0, 23.0, 89], 
             ["2Introduction to", 30.0, 26.0, 23.0, 89], 
             ["3Data Analytics", 30.0, 26.0, 23.0, 89], 
             ["4Data Analytics", 30.0, 26.0, 23.0, 89], 
             ["5Introduction to", 30.0, 26.0, 23.0, 89]]
        return a
    
    def table_summary(self, name, month):
        working_data  =  self.data
        return ["TOTAL", 89, 26.0, 23.0, None]
    
    def main_of_pdf(self, name, month):
        c  =  canvas.Canvas((self.file_loc + name+"_"+month+".pdf"), bottomup = 1, pagesize = A4)
        c = self.draw_border(c, 35)
        c = self.draw_border(c, 32.5)
        c = self.draw_border(c, 30)
        c = self.draw_intro(c, 25)
        c = self.draw_table(c)
        c = self.draw_comparison_table(c)
        c = self.draw_acknowledgement(c, 22)
        c.showPage()
        c.save()
    
    def draw_border(self, c, m):
        c.line(m, m, 595.27-m, m)
        c.line(m, 841.89-m, 595.27-m, 841.89-m)
        c.line(m, m, m, 841.89-m)
        c.line(595.27-m, m, 595.27-m, 841.89-m)
        return c
    
    def draw_intro(self, c, Spacing):
        c.setFont('Times-Bold', 28)
        c.setFillColorRGB(0, 0, 0.77)
        c.drawCentredString(595.27/2+50, 750, text = 'CampusX Mentorship Programme')
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Times-Roman', 22)
        c.drawCentredString(320, 720, text = 'Machine Learning')
        c.setFont('Times-Roman', 18)
        c.drawString(45, 680, ('NAME:-' + self.Name))
        c.drawString(45, 680-Spacing, 'COLLEGE:-'+self.College)
        # c.drawString(285, 680-2*Spacing, 'YEAR:-'+info.Year)
        c.drawString(45, 680-2*Spacing, 'MONTH:-'+self.Month)
        c.drawString(45, 680-3*Spacing, 'Email Address:- ' + self.Email)
        c.line(35, 680-3.5*Spacing, 560.27, 680-3.5*Spacing)
        c.drawInlineImage(image = "campusX_false_logo.png", x = 45, y = 700, width = 85, height = 100)
        c.drawInlineImage(image = "Photo_InkedcampusX_false_logo_LI.jpg", x = 440, y = 600, width = 100, height = 115)
        return c
    

    
    def draw_table(self, c):
        c.drawInlineImage(image = "TABLE_MODULES.jpg", x = 45, y = 410, width = 500, height = 180)
        c.setFont('Times-Bold', 10)
        Heading  =  ['MODULE', 'Full Marks', 'Heighest Marks', "Your Marks", 'Percentyle']
        for i in range(len(Heading)):
            c.drawCentredString(95+i*100, 575, Heading[i])
        writing_row  =  555
        data  =  self.Table_Content
        for i in range(len(self.Table_Content)):
            for j in range(len(self.Table_Content[i])):
                c.drawCentredString(95+j*100, writing_row, str(data[i][j]))   
            writing_row -=  20
        for i in range(len(self.Table_summary)):
            if(str(self.Table_summary[i])  ==  "None"):
                continue
            c.drawCentredString(95+i*100, 415, str(self.Table_summary[i])) 
        return c
    
    
    def draw_comparison_table(self, c):
        Y = 160     #Y scale of the second and third graph
        self.Give_me_first_graph_for_the_month_of(self.Name, self.Month)
        self.Creat_spided_plot(self.Name, self.Month)
        c.drawInlineImage(image = (self.file_loc + self.Name + "_" + self.Month + "_" + "graph1.jpg"), x = 45, y = Y, width = 290, height = 180)
        c.drawInlineImage(image = (self.file_loc + self.Name + "_" + self.Month + "Graph2.jpg"), x = 350, y = Y, width = 200, height = 180)
        c.setFont('Times-Roman', 19)
        c.drawString(45, 390, 'Comparison between this month and averagen till now on the ')
        c.drawString(55, 370, 'basic of:')
        c.drawString(65, 350, '1. Task Subject:')
        c.drawString(355, 350, '2. Task Value:')
        return c
    
    def Give_me_first_graph_for_the_month_of(self, Name, Month_name):
        working_data  =  self.data
        working_data  =  working_data[working_data["Student"] == Name]
        data_for_graph_one  =  working_data.groupby("Module")['Points', 'Total'].sum().reset_index()
        data_for_graph_one["Percentage"]  =  round(data_for_graph_one["Points"] / data_for_graph_one['Total']*100)
        data_for_graph_one["For_the_month_of"]  =  "All"
        data_for_given_month  =  working_data[working_data["Month"]  ==  Month_name]
        data_for_given_month  =  data_for_given_month.groupby("Module")['Points', 'Total'].sum().reset_index()
        data_for_given_month["Percentage"]  =  round(data_for_given_month["Points"] / data_for_given_month['Total']*100)
        data_for_given_month["For_the_month_of"]  =  str(Month_name)
        out_put_dataframe  =  pd.concat([data_for_given_month, data_for_graph_one])
        out_put_dataframe  =  out_put_dataframe[out_put_dataframe["Module"] !=  "Ritual"]
        plot  =  sns.barplot(x = 'Percentage', y = 'Module', data  =  out_put_dataframe, hue  =  'For_the_month_of', color  =  "Green", )
        plot.get_figure().savefig((self.file_loc + Name + "_" + Month_name + "_" + "graph1.jpg"), dpi = 1200, bbox_inches  =  'tight')
        
        
    def Creat_spided_plot(self, name, month):
        
        working_data  =  self.data
        working_data  =  working_data[working_data["Student"]  ==  name]
        working_data  =  working_data[working_data["Month"]  ==  month]
        My_list  =  []
        7


        for i in range(len(working_data["id"])):
            current_row  =  working_data.iloc[i].values
            jata  =  [current_row[4], current_row[7], current_row[8]]
            My_list.append(jata)
        transfer  =  My_list 
        next_my_list  =  []
        for i in transfer:
            a  =  i
            w  =  a[0].replace(" ", "").split(",")
            for j in w:
                one  =  j
                two  =  a[1]
                three  =  a[2]
                four  =  [one, two, three]
                next_my_list.append(four)
        spider_plot_df  =  pd.DataFrame(next_my_list)
        spider_plot_df.rename(columns = {0:"Type", 1:"Points", 2:"FM"}, inplace = True)
        marks  =  spider_plot_df.groupby("Type")["Points", "FM"].sum().reset_index()
        a  =  round(marks["Points"] / marks["FM"] * 100).values
        marks["Percentage"]  =  a 
        marks.index  =  marks["Type"]
        marks.drop(columns = {"Type"}, axis  =  1, inplace  =  True)
        # marks =  marks[marks["FM"]>marks["FM"].mean()]
        labels = np.array(marks["Percentage"].index[:])
        stats  =  marks["Percentage"].values[:]
        angles = np.linspace(0,  2*np.pi,  len(labels),  endpoint = False)
        # close the plot
        stats = np.concatenate((stats, [stats[0]]))
        angles = np.concatenate((angles, [angles[0]]))
        fig = plt.figure()
        ax  =  fig.add_subplot(111,  polar = True)
        ax.plot(angles,  stats,  'o-',  linewidth = 2)
        ax.fill(angles,  stats,  alpha = 0.25)
        ax.set_thetagrids(angles * 180/np.pi,  labels)
        # ax.set_title("Shakib")
        ax.grid(True)
        ax.get_figure().savefig((self.file_loc + name + "_" + month + "Graph2.jpg"), dpi = 1200, bbox_inches  =  'tight')
    
    def draw_acknowledgement(self, c, Spacing):
        c.line(35, 45+5*Spacing, 560.27, 45+5*Spacing)
    
        c.setFont('Times-Roman', 18)
        c.drawString(45,  45+4*Spacing,  'Number of task wins:-'+self.Number_of_task_wins)
        c.drawString(45, 45+3*Spacing, 'Rank among the class:-'+ self.Rank_among_the_class)
        c.drawString(45, 45+2*Spacing, 'Late submition ratio:-'+ self.Late_submition_ratio)
        c.drawString(45, 45+Spacing, 'Teacher’s signature:-')
        c.drawString(45, 45, 'Remark:-')
    
    
        c.drawString(300, 45+4*Spacing, 'Overall percentage :-'+ self.Percentage + "%")
        c.drawString(300, 45+3*Spacing, 'Overall1 percentyle:-'+ self.Percentile + "%")
        c.drawString(300, 45+2*Spacing, 'Generated on:-'+self.Date)
        return c
    
    

    
from datetime import date
from datetime import datetime
from reportlab.pdfgen import canvas
#from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
import pandas as pd
import numpy as np
import openpyxl
import requests
import seaborn as sns
import matplotlib.pyplot as plt
#from math import pi
import os
Shakib  =  Main()

