from viktor import ViktorController
from viktor.parametrization import ViktorParametrization, DownloadButton, NumberField, Table, TextField, DateField , Text, LineBreak
from viktor.result import DownloadResult
from viktor.external.word import WordFileComponent, WordFileTag, WordFileTemplate, render_word_file, WordFileImage
from viktor.views import ImageView, ImageResult, PDFView, PDFResult
from viktor.utils import convert_word_to_pdf
import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np 
from pathlib import Path
from io import StringIO

def generate_word_document(client_name,
                           adress,
                           zip_code,
                           city,
                           invoice_number,
                           date,
                           df_price,
                           total_price
                           ):
        components = []
        # Text tags 
        components.append(WordFileTag("Client_name", client_name))
        components.append(WordFileTag("address", adress))
        components.append(WordFileTag("zip_code", zip_code))
        components.append(WordFileTag("city", city))
        components.append(WordFileTag("date", date))
        components.append(WordFileTag("invoice_number", invoice_number))
        components.append(WordFileTag("total_price", str(total_price)))

        # Dynamic tables
        df_price = df_price.to_dict("records")
        components.append(WordFileTag('table1', df_price))
        components.append(WordFileTag("table2", df_price))

        # Figure 
        figure_path = Path(__file__).parent / "files" / "figure.png"
        with open(figure_path, "rb") as figure:
            word_file_figure = WordFileImage(figure, "figure_sales", width = 500)

        components.append(word_file_figure)

        template_path = Path(__file__).parent /"files" / "Template.docx"
        with open(template_path, 'rb') as template:
            word_file = render_word_file(template, components)
        
        return word_file

def create_figure(df_price):
    print(df_price)
    
    # Create figure
    fig, ax = plt.subplots(figsize=(16, 8))
    products = list(df_price["desc"])
    qty = list(np.round(df_price["qty"],2))
    ax.pie(qty, labels = products, autopct = "%1.1f%%")
    ax.set_title("Pie chart total sold products")
    plt.savefig("files/figure.png")
    plt.close()
    
    return fig


class Parametrization(ViktorParametrization):
    txt1 = Text("Welcome to this app, in this app it is possible to make an invoice")
    # Adding fields 
    client_name = TextField("Name of client")
    lb1 = LineBreak()
    adress = TextField("Adress")
    lb2 = LineBreak()
    zip_code = TextField("Zip code")
    lb3 = LineBreak()
    city = TextField("City")
    lb4 = LineBreak()
    invoice_number = TextField("invoice number")
    lb5 = LineBreak()
    date = DateField("Data")

    # Table
    table_price = Table("Products")
    table_price.qty = NumberField("Quantity", suffix="-")
    table_price.desc = TextField("Description", suffix="-")
    table_price.price = NumberField("Price", suffix="€")

    download_word_file = DownloadButton('Download report', method='download_word_file')
    
class Controller(ViktorController):
    viktor_enforce_field_constraints = True 

    label = 'My Entity Type'
    parametrization = Parametrization

    @ImageView("Plot", duration_guess=1)
    def visualize_image(self, params, **kwargs):
        df_price = pd.DataFrame.from_dict(params.table_price) 
        fig = create_figure(df_price)
        # Saving figure
        svg_data = StringIO()
        fig.savefig(svg_data, format='svg')

        return ImageResult(svg_data)

    def download_word_file(self, params, **kwargs):
        # Generating a panda dataframe from the price table 
        df_price = pd.DataFrame.from_dict(params.table_price) 

        total = []
        for i in df_price.index:
            qty = df_price.loc[i, "qty"]
            price = df_price.loc[i, "price"]
            total.append(qty*price)
        
        total_price = np.sum(total) # Calculating total price
        
        df_price["total"] = total # Adding total price column to dataframe

        perc_total = []
        for i in df_price.index:
            total = df_price.loc[i, "total"]
            calc = round((total/ total_price)*100,2)
            perc_total.append(str(calc)+"%")
        df_price["perc"] = perc_total # adding percentage column to dataframe

        figure = create_figure(df_price)
        word_file = generate_word_document(client_name = params.client_name,
                                           adress = params.adress,
                                           zip_code = params.zip_code,
                                           city = params.city,
                                           date = str(params.date),
                                           invoice_number = params.invoice_number,
                                           df_price = df_price,
                                           total_price = total_price
                                           )

        return DownloadResult(word_file,"test.docx")

    @PDFView("PDF viewer", duration_guess=5)  
    def pdf_view(self, params, **kwargs):
        df_price = pd.DataFrame.from_dict(params.table_price) 

        total = []
        for i in df_price.index:
            qty = df_price.loc[i, "qty"]
            price = df_price.loc[i, "price"]
            total.append(qty*price)
        
        total_price = np.sum(total) # Calculating total price
        
        df_price["total"] = total # Adding total price column to dataframe

        perc_total = []
        for i in df_price.index:
            total = df_price.loc[i, "total"]
            calc = round((total/ total_price)*100,2)
            perc_total.append(str(calc)+"%")
        df_price["perc"] = perc_total # adding percentage column to dataframe

        figure = create_figure(df_price)
        word_file = generate_word_document(client_name = params.client_name,
                                           adress = params.adress,
                                           zip_code = params.zip_code,
                                           city = params.city,
                                           date = str(params.date),
                                           invoice_number = params.invoice_number,
                                           df_price = df_price,
                                           total_price = total_price
                                           )
                                           
        with word_file.open_binary() as f1:
            pdf_file = convert_word_to_pdf(f1)

    
        return PDFResult(file = pdf_file)
