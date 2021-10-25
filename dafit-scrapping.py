import requests
from bs4 import BeautifulSoup
import xlsxwriter

# criando nosso arquivo excel
workbook = xlsxwriter.Workbook('produtos.xlsx')
worksheet = workbook.add_worksheet()

# criando cabeçalho
bold = workbook.add_format({'bold':1})
worksheet.write(0,0,"Nº",bold)
worksheet.write(0,1,"Produto",bold)
worksheet.write(0,2,"Marca",bold)
worksheet.write(0,3,"Categoria",bold)
worksheet.write(0,4,"Promoção?",bold)
worksheet.write(0,5,"Preço Original",bold)
worksheet.write(0,6,"Preço Promocional",bold)

# definindo linha e coluna iniciais
row = 1
col = 0

#iterando sobre as 71 páginas
for i in range(71):
    # acessando o código da página i+1 ...
    r = requests.get('https://www.dafiti.com.br/roupas-masculinas/calcas-jeans/?page='+ str(i + 1))
    html = r.text
    soup = BeautifulSoup(html, 'html.parser')

    # listando os produtos
    productBoxDetail = soup.find_all("div", {"class": "product-box-detail"})

    # iterando sobre os produtos
    for product in productBoxDetail[1:]:
        #descobrindo se o produto está na promoção
        isPromocao = True if product.find("div", {"class": "is-special-price"}) else False

        #escrevendo no arquivo
        worksheet.write(row,col,row)
        worksheet.write(row,col+1,product.find("p",{"class": "product-box-title"}).text)
        worksheet.write(row,col+2,product.find("div",{"class": "product-box-brand"}).text)
        worksheet.write(row,col+3,"Calças Jeans")
        worksheet.write(row,col+4,str(isPromocao))
        worksheet.write(row,col+5,float(product.find("span",{"class": "product-box-price-from"}).text.replace("R$ ","").replace(".","").replace(",",".")))
        worksheet.write(row,col+6,float(product.find("span",{"class": "product-box-price-to"}).text.replace("R$ ","").replace(".","").replace(",",".")) if isPromocao else "-")
        row += 1
# fechando planilha
workbook.close()