from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CDN='https://cdn.shopify.com/s/files/1/0986/2430/7486/files/'
LOJA='https://www.conttsfitness.com.br/products/'

# peça, tamanho, cor, qtd, sku, handle, arquivo da foto
ITENS=[
 ('Legging empina bumbum abstract branca','G','Branco',3,'CTTS-LEGEMPBUMA-G','legging-empina-bumbum-abstract-branca-lgkk','1-416f44017f4c9592a817805080725201-1024-1024.webp?v=1785963271'),
 ('Legging empina bumbum cintura alta jacquard preta','G','Preto',2,'CTTS-LEGEMPBUMC-G-02','legging-empina-bumbum-cintura-alta-jacquard-preta-cawa','IMG-0277.png?v=1783358892'),
 ('Top regata elástico basic preto','G','Preto',1,'CTTS-TOPELSBASP-G','top-elstico-basic-preto-qg66','40-f3248108a67af8bf9c17683274986240-1024-1024_9ea61c36-5dd1-4a54-9eb7-1a4adb22ef9f.webp?v=1782268418'),
 ('Top regata elástico basic marrom','G','Marrom',1,'CTTS-TOPELSBASM-G','top-elstico-basic-marrom-p8fy','2-60c2b070ce19f23fed17683268542879-1024-1024.webp?v=1782268420'),
 ('Legging cintura alta canelada capuccino milano','M','Capuccino',1,'CTTS-LEGCINMEDC-M-CAPU','legging-cintura-media-canelada-capuccino','IMG-0418.png?v=1785505880'),
 ('Legging empina bumbum grafiatto rosa','G','Rosa',1,'CTTS-LEGEMPBUMG-G-02','legging-empina-bumbum-grafiatto-rosa-vfvl','33-8fd1743dcfea31df6817707550159677-1024-1024_51e600a7-593a-4949-8b65-e8152787f19d.webp?v=1782268402'),
 ('Top milano canelado azul','G','Azul',1,'CTTS-TOPMILCANA-G','top-milano-canelado-azul-wc6c','31-141822f75843577e7117679019782975-1024-1024_02db8ad5-cdcd-4cd2-b817-f60a161440a8.webp?v=1782268629'),
 ('Legging empina bumbum fake jeans','G','Jeans',1,'CTTS-LEGEMPBUMF-G-02','legging-empina-bumbum-fake-jeans-zbin','2-5c83042612daaf810c17805080467916-1024-1024_5a944d60-d9fc-4cc6-8315-4ef73b31d9ef.webp?v=1782268518'),
 ('Legging comum marrom','P','Marrom',1,'CTTS-LEGCOMMAR-P','legging-comum-marrom-5wdp','45-8a7a0e0f3955c0243117707538654867-1024-1024_2a6cba1e-e527-4850-a7f1-1d86d2f043aa.webp?v=1782268516'),
 ('Macacão empina bumbum jacquard brocado Preto','Tamanho Único (38 a 42)','Preto',1,'CTTS-MACEMPBUMJ-UNI-PRET-02','macacao-empina-bumbum-jacquard-brocado-preto','8_4486ddb1-a7db-4cd7-aaba-20be4afdcbe1.png?v=1782930139'),
 ('Top regata abstract','G','Estampado',1,'CTTS-TOPREGABS-G','top-regata-abstract-b2jr','IMG-0398.png?v=1784317582'),
 ('Conjunto legging empina bumbum e Top urban ink','Top G / Legging G','Preto',1,'CTTS-CONLEGEMPB-URBINK-GxG','conjunto-legging-empina-bumbum-e-top-urban-ink','11_9804ea76-0238-4276-bcbd-6c1085170514.png?v=1783216409'),
 ('Conjunto Legging cintura alta modeladora brocada Azul Marinho','M','Azul Marinho',1,'CTTS-LEGCAMOD-M-MARI','conjunto-legging-cintura-alta-modeladora-brocada-azul-marinho','BD969A73-690E-4C05-9E0F-84C70299CC84.png?v=1785962921'),
 ('Top tulipa branco','M','Branco',1,'CTTS-TOPTULBRA-M','top-tulipa-branco-avsm','40-e8b91ecfe083067aa517680554163924-1024-1024.webp?v=1782268537'),
]

# pedido, data, cliente, valor, situação, [(peça, tamanho, cor, qtd, sku, idx do item)]
PEDIDOS=[
 ('#1072','29/08/2026','Maria Aparecida Pianezzer',711.55,'Enviar',[2,3,4,5,1]),
 ('#1073','30/08/2026','Magda Gomes',213.10,'Enviar',[6,7]),
 ('#1074','02/09/2026','Vanderlei Lucas',166.82,'Enviar',[8]),
 ('#1075','03/09/2026','Adriana FIALA',242.10,'NÃO ENVIAR - cancelado, duplicata do #1076',[9]),
 ('#1076','03/09/2026','Adriana FIALA',255.55,'Enviar',[9]),
 ('#1077','04/09/2026','Fabiana Pereira de Jesus Silveira',186.10,'Enviar',[0]),
 ('#1078','06/09/2026','Maiara Dias',583.30,'Enviar',[10,0,11]),
 ('#1079','07/09/2026','Roberto França Hid',205.05,'Enviar',[12]),
 ('#1080','07/09/2026','Julia Almeida',486.40,'Enviar',[13,1,0]),
]

F='Arial'
HDR=Font(name=F,bold=True,color='FFFFFF',size=10)
FILL=PatternFill('solid',fgColor='211E1A')
BODY=Font(name=F,size=10)
LINK=Font(name=F,size=10,color='0563C1',underline='single')
BOLD=Font(name=F,bold=True,size=10)
ALERT=Font(name=F,size=10,bold=True,color='9C0006')
ALERTF=PatternFill('solid',fgColor='FFC7CE')
THIN=Side(style='thin',color='D9D1C3')
BOX=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
WRAP=Alignment(vertical='center',wrap_text=True)
CEN=Alignment(horizontal='center',vertical='center')

wb=Workbook()

# ---------- aba 1: separar ----------
ws=wb.active; ws.title='Separar'
cab=['Separado','Peça','Tamanho','Cor','Qtd','SKU','Ver foto','Página na loja']
ws.append(cab)
for c in range(1,len(cab)+1):
    cel=ws.cell(row=1,column=c); cel.font=HDR; cel.fill=FILL; cel.alignment=CEN; cel.border=BOX
for i,(nome,tam,cor,qtd,sku,handle,foto) in enumerate(ITENS,start=2):
    ws.cell(row=i,column=1,value='').border=BOX
    ws.cell(row=i,column=2,value=nome).font=BODY
    ws.cell(row=i,column=3,value=tam).font=BOLD
    ws.cell(row=i,column=4,value=cor).font=BODY
    ws.cell(row=i,column=5,value=qtd).font=BOLD
    ws.cell(row=i,column=6,value=sku).font=Font(name=F,size=9,color='655E51')
    f=ws.cell(row=i,column=7,value='abrir foto'); f.hyperlink=CDN+foto; f.font=LINK
    l=ws.cell(row=i,column=8,value='abrir página'); l.hyperlink=LOJA+handle; l.font=LINK
    for c in range(1,9):
        ws.cell(row=i,column=c).border=BOX
        ws.cell(row=i,column=c).alignment=CEN if c in (1,3,5,7,8) else WRAP
tot=len(ITENS)+2
ws.cell(row=tot,column=4,value='TOTAL DE PEÇAS').font=BOLD
t=ws.cell(row=tot,column=5,value='=SUM(E2:E%d)'%(len(ITENS)+1)); t.font=BOLD; t.alignment=CEN
for w,c in zip([11,46,20,14,7,30,12,15],'ABCDEFGH'): ws.column_dimensions[c].width=w
for r in range(2,len(ITENS)+2): ws.row_dimensions[r].height=30
ws.freeze_panes='A2'
ws.auto_filter.ref='A1:H%d'%(len(ITENS)+1)

n=tot+2
ws.cell(row=n,column=2,value='Marque a coluna "Separado" conforme for tirando do estoque.').font=Font(name=F,size=9,italic=True,color='655E51')
ws.cell(row=n+1,column=2,value='"Ver foto" abre só a imagem. "Página na loja" abre o produto inteiro no site.').font=Font(name=F,size=9,italic=True,color='655E51')
a=ws.cell(row=n+2,column=2,value='ATENÇÃO: no conjunto urban ink a foto principal cadastrada na loja é a tabela de medidas. O link aqui aponta para a foto real da peça, que é a segunda imagem.')
a.font=Font(name=F,size=9,italic=True,color='9C0006')

# ---------- aba 2: por pedido ----------
w2=wb.create_sheet('Por pedido')
cab2=['Pedido','Data','Cliente','Situação','Peça','Tamanho','Cor','Qtd','SKU','Ver foto','Valor do pedido']
w2.append(cab2)
for c in range(1,len(cab2)+1):
    cel=w2.cell(row=1,column=c); cel.font=HDR; cel.fill=FILL; cel.alignment=CEN; cel.border=BOX
r=2
for num,data,cli,valor,sit,idxs in PEDIDOS:
    off = sit.startswith('NÃO')
    primeiro=True
    for k in idxs:
        nome,tam,cor,qtd,sku,handle,foto=ITENS[k]
        w2.cell(row=r,column=1,value=num).font=BOLD
        w2.cell(row=r,column=2,value=data).font=BODY
        w2.cell(row=r,column=3,value=cli).font=BODY
        sc=w2.cell(row=r,column=4,value=sit); sc.font=ALERT if off else BODY
        if off: sc.fill=ALERTF
        w2.cell(row=r,column=5,value=nome).font=BODY
        w2.cell(row=r,column=6,value=tam).font=BOLD
        w2.cell(row=r,column=7,value=cor).font=BODY
        w2.cell(row=r,column=8,value=1).font=BODY
        w2.cell(row=r,column=9,value=sku).font=Font(name=F,size=9,color='655E51')
        f=w2.cell(row=r,column=10,value='abrir foto'); f.hyperlink=CDN+foto; f.font=LINK
        if primeiro:
            v=w2.cell(row=r,column=11,value=valor); v.number_format='R$ #,##0.00'; v.font=BODY
            primeiro=False
        for c in range(1,12):
            w2.cell(row=r,column=c).border=BOX
            w2.cell(row=r,column=c).alignment=CEN if c in (1,2,6,8,10) else WRAP
        r+=1
fim=r+1
w2.cell(row=fim,column=5,value='PEÇAS A ENVIAR (sem o #1075)').font=BOLD
q=w2.cell(row=fim,column=8,value='=SUMIF(D2:D%d,"Enviar",H2:H%d)'%(r-1,r-1)); q.font=BOLD; q.alignment=CEN
w2.cell(row=fim+1,column=5,value='TOTAL DOS PEDIDOS A ENVIAR').font=BOLD
tv=w2.cell(row=fim+1,column=11,value='=SUMIF(D2:D%d,"Enviar",K2:K%d)'%(r-1,r-1)); tv.font=BOLD; tv.number_format='R$ #,##0.00'
for w,c in zip([10,12,32,34,42,18,14,7,30,12,16],'ABCDEFGHIJK'): w2.column_dimensions[c].width=w
for rr in range(2,r): w2.row_dimensions[rr].height=28
w2.freeze_panes='A2'
w2.auto_filter.ref='A1:K%d'%(r-1)
w2.cell(row=fim+3,column=5,value='Nenhum destes pedidos tem baixa de envio na Shopify — as clientes ainda não receberam e-mail de postagem.').font=Font(name=F,size=9,italic=True,color='655E51')

wb.save('separacao-pedidos.xlsx')
print('planilha gerada')
