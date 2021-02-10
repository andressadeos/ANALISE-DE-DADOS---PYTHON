import pandas as pd
#Passo 01 - Importa banco de dados
df_tabela_vendas=pd.read_excel(r"C:\Andressa\workspace\VENDAS_REL\vendas.xlsx", engine='openpyxl')
print(df_tabela_vendas)
#Passo 02 - Visualizar a Base de Dados para ver se precisamos fazer algum tratamento
print(df_tabela_vendas.info())
#Passo 03 - Calcular os indicadores de todas as lojas:
# Faturamento
df_tabela_faturamento = df_tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
df_tabela_faturamento = df_tabela_faturamento.sort_values(by='Valor Final', ascending=False)
print(df_tabela_faturamento)
#Quantidade de Produto por lojas
df_tabela_produto = df_tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
df_tabela_produto = df_tabela_produto.sort_values(by='Quantidade', ascending=False)
print(df_tabela_produto)
#Quantidade de Produto por lojas
df_tabela_ticket_medio = (df_tabela_faturamento['Valor Final'] / df_tabela_produto['Quantidade']).to_frame()
df_tabela_ticket_medio = df_tabela_ticket_medio.rename(columns={0: 'Ticket Medio'})
df_tabela_ticket_medio = df_tabela_ticket_medio.sort_values(by='Ticket Medio', ascending=False)
print(df_tabela_ticket_medio)
#função de enviar e-mail
def enviar_email(loja,resumo_loja):
    import smtplib
    import email.message

    server = smtplib.SMTP('smtp.gmail.com:587')  
    corpo_email = f"""
    <p>Prezados,</p>
    <p>Segue relatorio de Vendas</p>
    {resumo_loja.to_html()}
    <p>Qualquer duvida estou a disposição!</p>
    <p>Att</p>
    """
    
    msg = email.message.Message()
    msg['Subject'] = f"Relatorio de Vendas - {loja}"
    
    # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
    # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
    # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
    msg['From'] = 'remetente@gmail.com'
    msg['To'] = 'destinatario@gmail.com'
    password = "senha"
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )
    
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('E-mail enviado')

#Passo 04 - Calcular os indicadores de cada loja
lista_lojas=df_tabela_vendas['ID Loja'].unique()

for loja in lista_lojas: 
    df_tabela_loja= df_tabela_vendas.loc[df_tabela_vendas['ID Loja']==loja, ['ID Loja', 'Quantidade', 'Valor Final']]#.loc filtra todas as linhas das colunas ID Loja, quantidade e valor final de cada loja
    df_tabela_loja = df_tabela_loja.groupby('ID Loja').sum()
    df_tabela_loja['Ticket Medio']=df_tabela_loja['Valor Final'] / df_tabela_loja['Quantidade']#criando uma columa dentro da tabela
#Passo 5 - Enviar e-mail para cada loja
    enviar_email(loja, df_tabela_loja)
#Passo Passo 6 - Enviar e-mail para a diretoria
tabela_completa = df_tabela_faturamento.join(df_tabela_produto).join(df_tabela_ticket_medio)
enviar_email("Diretoria", tabela_completa)