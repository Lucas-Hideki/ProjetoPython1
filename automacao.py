import pandas as pd
import win32com.client as win32
import pathlib

emails = pd.read_excel(r'C:\Users\Master\Documents\Programas\Caderno\Projetos\Projeto 1\Emails.xlsx')
lojas = pd.read_excel(r'C:\Users\Master\Documents\Programas\Caderno\Projetos\Projeto 1\Lojas.xlsx')
vendas= pd.read_excel(r'C:\Users\Master\Documents\Programas\Caderno\Projetos\Projeto 1\Vendas.xlsx')

#print(emails, lojas, vendas)

# analisando as planilhas de vendas e lojas, nota-se que a ambas as planilhas possuem um ID, porém a planilha de vendas não tem uma coluna com as lojas. Com isso, vou mesclar a planilha de vendas com a planilha de lojas para puxar esssa coluna. (para mesclar uma planilha com outra, usamos a váriavel.merge)
vendas = vendas.merge(lojas) #usamos o on, para dizer o local onde será mesclado
#display(vendas)

#agora, precisamos de um for para percorrer a planilha de vendas com o objetivo de criar uma tabela para cada loja, para isso, utilizarei o for. Para armazenar cada valor, vou criar um dicionário fora do for
dicionario_lojas = {}
# o for basicamente diz o seguinte: para cada item(loja) dentro da minha variável lojas, eu quero percorrer a minha coluna Loja
for loja in lojas["Loja"]:
    #usamos o [loja] para adicionar cada item da coluna no dicionário. Vamos filtrar a planilha de vendas com o .loc[linhas, colunas], onde a coluna será a mesma, ou seja, basta usar o ":"; e a linha será a mesma linha da loja
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"]==loja, :]

# agora que ja criei as tabelas, vou criar os indicadores de dia e ano na tabela. Os indicadores vão ter que calcular, mesmo quando entrar dias mais recentes. Para pegar o dia atual, basta usar o .max

dia_indicador = vendas["Data"].max()
print(dia_indicador)
print(f"{dia_indicador.day}/{dia_indicador.month}") #podemos formatar ainda para mostrarmos a data da forma que quisermos

#Agora, vou mexer com a pasta de backup. Primeiro vou verificar se uma pasta ja existe e caso não exista, vou criar uma pasta nova. Para criar as pastas, vou usar a biblioteca pathlib

caminho_pasta = pathlib.Path(r"backup arquivos lojas")
# agora vou usar o método iterdir que serve para pegar todos os objetos existentes dentro da pasta
arquivos_pasta = caminho_pasta.iterdir()
#agora vou usar o for no arquivos_pasta para percorrer minha pasta do pc. Para adicionar um arquivo, usamos o append
#lista_nome = [arquivo.name for arquivo in arquivos_pasta] #uma forma alternativa seria usando o list comprehension

lista_nome = []

for arquivo in arquivos_pasta:
    lista_nome.append(arquivo.name)
print(lista_nome)

# agora vou criar as pastas com os nomes de cada uma das lojas. Para isso vou percorrer o dicionário criado anteriormente com cada loja e verificar se essa loja existe dentro da lista_nome. Se existir, vou apenas passar para o próximo item e se não existir vou criar um novo com o nome da loja

for loja in dicionario_lojas:
    if loja not in lista_nome: #se não estiver dentro da pasta, criarei um novo com o método mkdir(). Podemos usar a própria variável com o caminho da pasta e concatenar com o nome do arquivo: (isso é possivel porque a biblioteca pathlib permite isso)
        nova_pasta = caminho_pasta / loja
        nova_pasta.mkdir()
    #agora que criei e vefiriquei todas as pastas, vou mudar o nome do arquivo para ficar formatado com a data e com a extensão do arquivo
    nome_arquivo = (f"{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx")
    local_arquivo = caminho_pasta / loja / nome_arquivo
    #agora, vou criar um novo excel para cada tabela criada com o dicionario lojas
    dicionario_lojas[loja].to_excel(local_arquivo)

#Definindo as metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdprodutos_dia = 4
meta_qtdprodutos_ano = 120

# Os indicadores que precisam ser calculados são: Faturamento (dia e ano) e produtos 

loja = "Norte Shopping"
vendas_loja = dicionario_lojas[loja] 
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

#faturamento
faturamento_anual = vendas_loja["Valor Final"].sum()
#print(faturamento_anual)
faturamento_diario = vendas_loja_dia["Valor Final"].sum()
#print(faturamento_diario)

#produtos
qtde_produtos_ano = len(vendas_loja["Produto"].unique()) #Para pegar os itens exclusivos
#print(qtde_produtos_ano)
qtde_produtos_diario = len(vendas_loja_dia["Produto"].unique())
#print(qtde_produtos_diario)

def enviar_email():
    for loja in dicionario_lojas:

        vendas_loja = dicionario_lojas[loja] 
        vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

        #faturamento
        faturamento_anual = vendas_loja["Valor Final"].sum()
        faturamento_diario = vendas_loja_dia["Valor Final"].sum()

        #produtos
        qtde_produtos_ano = len(vendas_loja["Produto"].unique()) #Para pegar os itens exclusivos
        qtde_produtos_diario = len(vendas_loja_dia["Produto"].unique())
        outlook = win32.Dispatch('outlook.application')

        nome = emails.loc[emails["Loja"]==loja, "Gerente"].values[0] #pegando os nomes e emais na planilha emails em cada loja 
        mail = outlook.CreateItem(0)
        mail.To = emails.loc[emails["Loja"]==loja, "E-mail"].values[0] #O .values é importante para trazer apenas o valor do nome e email. Sem ele, o valor sai em formato de lista com o indice e o nome
        mail.Subject = f'OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
        #mail.Body = 'Texto do email'

        #criando as variáveis para usar no corpo do email
        if faturamento_diario >= meta_faturamento_dia:
            cor_fat_dia = "green"
        else:
            cor_fat_dia = "red"

        if faturamento_anual >= meta_faturamento_ano:
            cor_fat_ano = "green"
        else:
            cor_fat_ano = "red"

        if qtde_produtos_diario >= meta_qtdprodutos_dia:
            cor_prod_dia = "green"
        else:
            cor_prod_dia = "red"

        if qtde_produtos_ano >= qtde_produtos_ano:
            cor_prod_ano = "green"
        else:
            cor_prod_ano = "red"

        #Formatando o corpo do email com html
        mail.HTMLBody = f''' 
        <p>Bom dia , {nome} </p>
        <p>O Resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:

        <table> 
            <tr>
                <th>Indicador</th>
                <th>Valor Dia</th>
                <th>Meta Dia</th>
                <th>Cenário Dia</th>
            </tr>
            <tr>
                <td>Faturamento</td>
                <td style="text-align:center">R${faturamento_diario:.2f}</td>
                <td style="text-align:center">R${meta_faturamento_dia:.2f}</td>
                <td style="text-align:center"><font color="{cor_fat_dia}">◙</td>
            </tr>
            <tr>
                <td>Diversidade de Produtos</td>
                <td style="text-align:center">{qtde_produtos_diario}</td>
                <td style="text-align:center">{meta_qtdprodutos_dia}</td>
                <td style="text-align:center"><font color="{cor_prod_dia}">◙</td>
            </tr>
        </table>
        <br>
        <table> 
            <tr>
                <th>Indicador</th>
                <th>Valor Ano</th>
                <th>Meta Ano</th>
                <th>Cenário Ano</th>
            </tr>
            <tr>
                <td>Faturamento</td>
                <td style="text-align:center">R${faturamento_anual:.2f}</td>
                <td style="text-align:center">R${meta_faturamento_ano:.2f}</td>
                <td style="text-align:center"><font color="{cor_fat_ano}">◙</td>
            </tr>
            <tr>
                <td>Diversidade de Produtos</td>
                <td style="text-align:center">{qtde_produtos_ano}</td>
                <td style="text-align:center">{meta_qtdprodutos_ano}</td>
                <td style="text-align:center"><font color="{cor_prod_ano}">◙</td>
            </tr>
        </table>
        <p>Segue em anexo a planilha com todos os dados para mais detalhes </p>

        <p>Qualquer dúvida estou à disposição.</p>
        Att., Lucas
        '''

        attachment = pathlib.Path.cwd() / caminho_pasta / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
        #CWD basicamente pega o local do arquivo onde está o código 
        mail.attachments.Add(str(attachment)) #esse str é importante para que a concatenação da biblioteca pathlib funcione

        mail.Send()
        print(f'Email da loja {loja} enviado!')

def ranking_email():
    faturamento_lojas_ano = vendas.groupby("Loja")[["Loja", "Valor Final"]].sum()
    faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by="Valor Final", ascending=False) #para pegar os valores do maior para o menor

    nome_arquivo = f"{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx"
    faturamento_lojas_ano.to_excel(r"backup arquivos lojas\{}".format(nome_arquivo))

    vendas_dia = vendas.loc[vendas["Data"]==dia_indicador, :]
    faturamento_lojas_dia = vendas_dia.groupby("Loja")[["Loja", "Valor Final"]].sum()
    faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by="Valor Final", ascending=False)

    nome_arquivo = f"{dia_indicador.month}_{dia_indicador.day}_Ranking dia.xlsx"
    faturamento_lojas_dia.to_excel(r"backup arquivos lojas\{}".format(nome_arquivo))

    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails["Loja"]=="Diretoria", "E-mail"].values[0]
    mail.Subject = f"Ranking Dia {dia_indicador.day}/{dia_indicador.month}"
    mail.Body = '''
Prezados, bom dia

Melhor loja do dia em faturamento: Loja {} com faturamento R${:.2f}
Pior loja do dia em faturamento: Loja {} com faturamento R${:.2f}

Melhor loja do ano em faturamento: Loja {} com faturamento R${:.2f}
Pior loja do ano em faturamento: Loja {} com faturamento R${:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição

Att.,
Lucas
'''.format(
    faturamento_lojas_dia.index[0],
    faturamento_lojas_dia.iloc[0, 0],
    faturamento_lojas_dia.index[-1],
    faturamento_lojas_dia.iloc[-1, 0],
    faturamento_lojas_ano.index[0],
    faturamento_lojas_ano.iloc[0, 0],
    faturamento_lojas_ano.index[-1],
    faturamento_lojas_ano.iloc[-1, 0]
)

    attachment = pathlib.Path.cwd() / caminho_pasta / f"{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx"
    mail.Attachments.Add(str(attachment))
    attachment = pathlib.Path.cwd() / caminho_pasta / f"{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx"
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print("E-mail da diretoria enviado ")

enviar_email()
ranking_email()