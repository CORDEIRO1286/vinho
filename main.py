mport openpyxl as op

wb=op.open('FOLLOW-UP - TRANSANCHES.xlsx')

ws=wb.worksheets

notas_vinho=[]

data_emissao_Nf=[]

notas_transanches=[]

wb1=op.open('controleTransanches.xlsx')
ws1=wb1.worksheets
notasNãoecontrada=[]
oco={}
notas=[]
for j in range(2,ws1[0].max_row):
    if ws1[0].cell(row=j,column=4).value=='Entrega realizada normalmente':
        ws1[0].cell(row=j,column=4,value='FINALIZADA')
    elif ws1[0].cell(row=j,column=4).value=="Processo de transporte já iniciado":
         ws1[0].cell(row=j,column=4,value="Aguardando a Descarga/ transito")
    nota=str(ws1[0].cell(row=j,column=9).value)
    if "," in nota:
        nota=nota.split(',')
        for n in nota:
            notas.append(n)
            oco.get(n)
            oco[n] = ws1[0].cell(row=j, column=4).value
    else:
        notas.append(nota)
        oco.get(nota)
        oco[nota] = ws1[0].cell(row=j,column=4).value

for i in range(3,ws[0].max_row):
    if ws[0].cell(row=i,column=1).value in notas:
       ws[0].cell(row=i,column=13,value=oco[nota])
    else:
        notas_vinho.append(ws[0].cell(row=i,column=1).value)

wb.save("controletransanches1.xlsx")
