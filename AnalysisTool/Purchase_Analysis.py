import pandas as pd
import numpy as np
import xlsxwriter
import tkinter as tk
from tkinter import messagebox
import os



def Purchase_Analysis (root,input_file_entry,process_button,progress_bar) :
    
    input_file = input_file_entry.get()
    process_button.config(state=tk.DISABLED)
    process_button.config(text="Processing...")


    try:
        messagebox.showinfo("Information", "Processing data. Please wait...")

       
        df= pd.read_excel(input_file, sheet_name='Mb51')
        #Artype = pd.read_excel(input_file, sheet_name='Mattype')
        df['Code mouvement'] = pd.to_numeric(df['Code mouvement'], errors='coerce')

        
        progress_bar['value'] = 0
        root.update_idletasks()  

        df = df[(df['Ordre'].isnull()) & (df['Code mouvement'].isin([101, 102]))]
        df.drop(columns=['Ordre'], inplace=True)

        grouped_df = df.groupby([df['Date comptable'].dt.to_period('M'), 'Article']).agg({
            'Quantité': 'sum',
            'Montant DI': 'sum'
        }).reset_index()
        
        progress_bar['value'] = 25
        root.update_idletasks()  

        article_to_designation = df.set_index('Article')['Désignation article'].to_dict()
        unité_to_article = df.set_index('Article')['UQ de saisie'].to_dict()
        #article_to_mattype = Artype.set_index('Article')['MatTyp'].to_dict()
        #grouped_df['MatType'] = grouped_df['Article'].map(article_to_mattype)
        grouped_df['Désignation article'] = grouped_df['Article'].map(article_to_designation)
        grouped_df['UQ de saisie'] = grouped_df['Article'].map(unité_to_article)

        
        progress_bar['value'] = 50
        root.update_idletasks()  

        grouped_df['Montant DI unitaire'] = np.where(grouped_df['Quantité'] == 0, 0, grouped_df['Montant DI'] / grouped_df['Quantité'])

        reorder = ['Article', 'Désignation article', 'Quantité', 'Montant DI','Montant DI unitaire','UQ de saisie', 'Date comptable']
        grouped_df = grouped_df[reorder]


        
        progress_bar['value'] = 100
        root.update_idletasks()  

        result_file_final = "Purchase_Analysis.xlsx"
        with pd.ExcelWriter(result_file_final, engine='xlsxwriter') as writer:
            grouped_df.to_excel(writer, sheet_name='Purchase Analysis', index=False)
            
                        

            workbook = writer.book
            worksheet_visualization = workbook.add_worksheet('Visualization')
                        
            

            chart_line_di = workbook.add_chart({'type': 'line'})
            chart_line_di.add_series({
                'categories': ['Purchase Analysis', 1, 6, len(grouped_df), 6],  
                'values':     ['Purchase Analysis', 1, 4, len(grouped_df), 4],  
                'name':       'Montant DI Unitaire ',
            })
            chart_line_di.set_size({'width': 600, 'height': 360})
            worksheet_visualization.insert_chart('M1', chart_line_di)


            chart_line_quantity = workbook.add_chart({'type': 'line'})
            chart_line_quantity.add_series({
                'categories': ['Purchase Analysis', 1, 6, len(grouped_df), 6],  
                'values':     ['Purchase Analysis', 1, 2, len(grouped_df), 2],  
                'name':       'Quantité',
            })
            chart_line_quantity.set_size({'width': 600, 'height': 360})
            worksheet_visualization.insert_chart('M20', chart_line_quantity)

            chart_bar_di = workbook.add_chart({'type': 'bar'})
            chart_bar_di.add_series({
            'categories': ['Purchase Analysis', 1, 6, len(grouped_df), 6],  
            'values':     ['Purchase Analysis', 1, 3, len(grouped_df), 3],  
            'name':       'Montant DI',
            })
            chart_bar_di.set_size({'width': 600, 'height': 360})
            worksheet_visualization.insert_chart('A1', chart_bar_di)

            chart_bar_di2 = workbook.add_chart({'type': 'bar'})
            chart_bar_di2.add_series({
            'categories': ['Purchase Analysis', 1, 6, len(grouped_df), 6],  
            'values':     ['Purchase Analysis', 1, 2, len(grouped_df), 2],  
            'name':       'Quantity',
            })
            chart_bar_di2.set_size({'width': 600, 'height': 360})
            worksheet_visualization.insert_chart('A20', chart_bar_di2)

        messagebox.showinfo("Information", "Data processing completed. Result file saved as {}".format(result_file_final))
    except Exception as e:
        tk.messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        
        process_button.config(state=tk.NORMAL)
        process_button.config(text="Process Data")
    progress_bar['value'] = 0
    root.update_idletasks()  

    os.startfile(result_file_final)