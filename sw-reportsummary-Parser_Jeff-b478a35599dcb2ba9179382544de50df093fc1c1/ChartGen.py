import pandas as pd
import matplotlib.pyplot as plt

def create_pie_chart_from_sheet(input_excel_file,input_sheet_name,export_cloumns,export_img_file):
    df = pd.read_excel(input_excel_file, sheet_name=input_sheet_name)
    df = df.drop(0) # Drop the merge cell
    df.columns = export_cloumns # set column name
    plot = df.groupby(['Card type'])['Pass'].count().plot(kind='pie',y='Pass', title="Card type test count", autopct='%1.0f%%')
    fig = plot.get_figure()
    fig.savefig(export_img_file)
    plot.cla()


def create_pie_chart_from_data(data,group,title,y_title,export_img_file):
    df = pd.DataFrame(data)
    # print(df)
    plot = df.groupby([group]).sum().plot(kind='pie', y=y_title,title=title, autopct="%1.1f%%")
    fig = plot.get_figure()
    fig.savefig(export_img_file)
    plot.cla()

def create_pie_chart_from_data_label(data,labels,title,export_img_file):
    plt.pie(data,labels=labels, autopct='%1.1f%%')
    plt.title(title)
    plt.axis('equal')
    plt.savefig(export_img_file,dpi=70)
    plt.clf()