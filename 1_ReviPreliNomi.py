# REVISIÓN PRELIMINAR DE NÓMINA

#----Librerías----------------------------------------------------------------#

from msilib import text 
from turtle import color
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import openpyxl
import string


#------Funciones--------------------------------------------------------------#

# Encontrar la columna en excel ingresando un valor numerico. La 0 es la A.
def num_a_col_excel(num):
    col = ""
    while num > 0:
        num -= 1
        col = chr(num % 26 + 65) + col
        num = num // 26
    return col
    
#FORMATEA ARCHIVO DE NOMINA HORIZONTAL
def format_nomina_horizontal(path_nom_hori):

    """
    Esta función toma como argumento el archivo de nómina horizontal tal y como se genera del software NgSoft
    y lo formatea de modo que sea trabajable (quita columnas compartidas y conserva solo información necesaria)
    
    """

    df=pd.read_excel(path_nom_hori,header=8)
    columns=df.columns.tolist()

    #Creamos Lista con nombres de columnas que no contienen "Unnamed (estas surgen por dejar de compartir una celda)"
    names=[]

    for column in columns:
        if "Unnamed" in column:
            pass
        else:
            temp=column
            names.append(temp)

    #Conservamos solo columnas que no contienen "Unnamed"
    df_nom=df[df.columns[df.columns.isin(names)]]#.dropna()

    #Cambiamos el nombre de algunas columnas por como aparece en la primera fila - omitimos cuando fila 1 = nan
    column_names=df_nom.columns.tolist()
    row_1=df_nom.iloc[0,:].tolist()

    test=pd.DataFrame(column_names,row_1).reset_index().rename(columns={0:"name","index":"row1"}).dropna()

    #Creamos el diccionario para renombrar
    dict_names=dict(zip(test.name, test.row1))

    #Renombrando columnas y quitando total
    df_nom=df_nom.rename(columns=dict_names).iloc[1:,:-3].reset_index(drop=True).dropna()

    #Eliminamos dígito de verificación a la columna CC:
    CEDULA=[]
    for item in df_nom.CC:
        if isinstance(item,float):
            temp=int(item)
            CEDULA.append(temp)
        elif isinstance(item,int):
            temp=item
            CEDULA.append(temp)
        elif '-' in item:
            temp=item.split("-",1)[0]
            CEDULA.append(temp)
        else:
            temp=item
            CEDULA.append(temp) 

    df_nom.CC=CEDULA
    df_nom["CC"]=df_nom["CC"].astype("int64")

    #Reemplazamos "," por nada para poder sumar valores en cada columna que aplique:
    for column in df_nom.columns.tolist():    
        if "-" in column:
            aux=[]
            for value in df_nom[column]:
                temp=value.replace(",","")
                temp=float(temp)
                aux.append(temp)
            df_nom[column]=aux
        else:
            pass

    return df_nom

#MAESTRO DE NÓMINA PARA FECHA DE INGRESO Y CODIGO
def maestro(path_maestro):

    """ 
    Esta función lee el informe MAESTRO que se genera de Ngsoft y 
    conserla las columnas que necesitamos: "CODIGO" y "FECHA DE INGRESO"

    """

    df_2=pd.read_excel(path_maestro)#,  engine='openpyxl')
    df_2.rename(columns={df_2.columns[3]:"CC"},inplace=True)
    df_2=df_2[["CC","codigo_empleado","fecha_ingreso_contrato"]]
    
    #Borrando los duplicados y conservando la fecha más reciente:  
    df_2=df_2.sort_values('fecha_ingreso_contrato').drop_duplicates('CC',keep='last')
    return df_2

#DISTRIBUCIÓN DE SALARIOS DEL MES ANTERIOR
def rev_nom_anterior(path_revi_nomi):

    """"
    Esta función toma la revisión del mes anterior para traer las columnas
    "TIPO DE SALARIO", "% AL 100", "% FLI SALARIO BASICO", "% FLI", "SALARIO TOTAL"
    """

    df_5=pd.read_excel(path_revi_nomi)
    df_5=df_5.iloc[:,[0,9,10,11,12,13]]
    return df_5

#INCREMENTOS SALARIALES
def novedades_nomina(path_novedades,df_revision):

    """
    Esta función abre el archivo de novedades de nómina para aplicar aumentos salariales
    
    """
    xls = pd.ExcelFile(path_novedades)

    #Leemos la hoja donde están los aumentos salariales (16):
    df14_16=xls.parse(16,header=2).iloc[:,[0,6]]

    
    df14_16.rename(columns={"CEDULA ":"CC",df14_16.columns[1]:"TOTAL SALARIO"},inplace=True)

    global CC_INCRE,df_replace
    #Creamos lista de de CEDULA y de TOTAL SALARIO:
    CC_INCRE=list(df14_16["CC"])

    #FILTAMOS EL ARCHIVO DE REVISIÓN PARA LAS PERSONAS QUE TUVIERON INCREMENTO PARA REEMPLAZAR EL VALOR:

    df_replace=df_revision[df_revision["CC"].isin(CC_INCRE)]
    df_no_replace=df_revision[df_revision["CC"].isin(CC_INCRE)==False]

    df_replace["SALARIO TOTAL"]=list(df_replace.merge(df14_16,on="CC",how="left")["TOTAL SALARIO"])
    
    df_replace["AUMENTO DE SALARIO"] = df_replace["CC"].apply(lambda x: "SI" if x in CC_INCRE else "NO")


    return df_replace, df_no_replace

#Leer info de personal que ingresa
def ingresos(path_novedades,df_revision):
    global fecha

    """
    Esta función extrae la información del archivo de novedades de las personas 
    que ingresaron en el mes de revisión a la compañía, las variables que extrae
    son: "COMPENSACION", "TIPO DE SALARIO", "FECHA DE INGRESO".

    """

    xls = pd.ExcelFile(path_novedades)    
    df14_6=xls.parse(5)

    
    df14_6=df14_6.iloc[:,[0,3,4,5]].dropna()
    df14_6[df14_6.columns[0]]=df14_6[df14_6.columns[0]].astype("int64")
    df14_6=df14_6[df14_6[df14_6.columns[3]]<=fecha].reset_index(drop=True)
    
    

    #Listas:
    CC_INGRE=list(df14_6[df14_6.columns[0]])
    

    #Excluimos Documento de extranjeria dado que no coincide con el archivo 1
    CC_AUX=[]
    for value in CC_INGRE:    
        if len(str(value))>=13:
            pass
        else:
            temp=value
            CC_AUX.append(temp)
    CC_INGRE=CC_AUX
    
    df14_6=df14_6[df14_6[df14_6.columns[0]].isin(CC_INGRE)].reset_index(drop=True)
    
    
    df14_6.rename(columns={df14_6.columns[0]:"CC"},inplace=True)
    df14_6.rename(columns={df14_6.columns[2]:"TIPO SALARIO"},inplace=True)

    #Personal a modificar:

    df_replace_2=df_revision.loc[df_revision["CC"].isin(CC_INGRE)].reset_index(drop=True)
    df_no_replace_2=df_revision[df_revision["CC"].isin(CC_INGRE)==False]


    df_replace_2["SALARIO TOTAL"]=list(df_replace_2.merge(df14_6,on="CC",how="left")["COMPENSACION "])
    df_replace_2["TIPO DE SALARIO "]=list(df_replace_2.merge(df14_6,on="CC",how="left")["TIPO SALARIO"])
    

    return df_replace_2,df_no_replace_2

# CONFIGURA LOS PORCENTAJES DE PARTICIPIACIÓN DEL SALARIO
def distribucion_salario(df_revision):

    """
    Esta función toma el archivo de revisión y calcula los porcentajes de salario básico, flexible 
    dependiendo del tpo de contrato y de salario
    
    """

    #Quitamos espacios a texto de columna TIPO SALARIO:
    df_revision["TIPO DE SALARIO "]=df_revision["TIPO DE SALARIO "].replace("PLENO","ORDINARIO")
    df_revision["TIPO DE SALARIO "]=df_revision["TIPO DE SALARIO "].str.strip()

    #Agregar % que hacen falta del personal nuevo:

    #cambios a salario flexible

    df_revision.loc[(df_revision["%AL 100"].isna()) & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),"%AL 100"]=1.0
    df_revision.loc[(df_revision["%SALARIO BASICO "].isna()) & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),"%SALARIO BASICO "]=0.66
    df_revision.loc[(df_revision["% FLI "].isna()) & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),'% FLI ']=0.34

    #cambios a salario ordinario

    df_revision.loc[(df_revision["%AL 100"].isna()) & (df_revision["TIPO DE SALARIO "]=="ORDINARIO"),"%AL 100"]=1.0
    df_revision.loc[(df_revision["%SALARIO BASICO "].isna()) & (df_revision["TIPO DE SALARIO "]=="ORDINARIO"),"%SALARIO BASICO "]=1.0
    df_revision.loc[(df_revision["% FLI "].isna()) & (df_revision["TIPO DE SALARIO "]=="ORDINARIO"),'% FLI ']=0.0

    #cambios a sostenimiento aprendiz

    df_revision.loc[(df_revision["%AL 100"].isna()) & (df_revision["TIPO DE SALARIO "]=="APOYO SOSTENIMIENTO"),"%AL 100"]=1.0
    df_revision.loc[(df_revision["%SALARIO BASICO "].isna()) & (df_revision["TIPO DE SALARIO "]=="APOYO SOSTENIMIENTO"),"%SALARIO BASICO "]=1.0
    df_revision.loc[(df_revision["% FLI "].isna()) & (df_revision["TIPO DE SALARIO "]=="APOYO SOSTENIMIENTO"),'% FLI ']=0.0

    #cambios a salario flexible integral

    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),"%AL 100"]=1.0
    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),"%SALARIO BASICO "]=0.9
    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLEXIBLE"),'% FLI ']=0.1

    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLI"),"%AL 100"]=1.0
    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLI"),"%SALARIO BASICO "]=0.9
    df_revision.loc[(df_revision["TIPO"]=="INTEGRAL") & (df_revision["TIPO DE SALARIO "]=="FLI"),'% FLI ']=0.1
    
    df_revision.loc[df_revision["%AL 100"]==0,"%AL 100"]=1.0


    return df_revision

#CREA ARCHIVO PARA OPERAR:
def crea_archivo_pruebas(path_archivo_base,df_revision):
    df_base=pd.read_excel(path_archivo_base)
    df_base["ARCHIVO"]="BASE"

    df_revision["ARCHIVO"]="REVISION"

    #Append al archivo ya organizado del nuevo archivo a revisar:
    df_pruebas=pd.concat([df_base,df_revision]).fillna(0)
    df_pruebas=df_pruebas[df_pruebas["ARCHIVO"]=="REVISION"].drop(columns="ARCHIVO")

    return df_pruebas

#Funciones MAIN
def compila_archivo(path_nom_hori,path_maestro,path_revi_nomi,path_novedades,path_archivo_base):

    """
    Esta función integra todas las funciones que realizan cambios y uniones de archivos y retorna \n
    un solo DataFrame con el formato final listo para realizar los cáldulos de conceptos. 
    """

    global df_revision

    df_revision=format_nomina_horizontal(path_nom_hori).merge(maestro(path_maestro),on="CC",how="left")
    df_revision.rename(columns={"codigo_empleado":"CODIGO","fecha_ingreso_contrato":"FECHA DE INGRESO"},inplace=True)
    
    
    
    df_revision=df_revision.merge(rev_nom_anterior(path_revi_nomi),on="CC",how="left")
    df_revision['SALARIO TOTAL']=df_revision['SALARIO TOTAL'].astype(float)
    
    
    
    df_revision=pd.concat([novedades_nomina(path_novedades,df_revision)[0],novedades_nomina(path_novedades,df_revision)[1]]).reset_index(drop=True)


    
    df_revision=pd.concat([ingresos(path_novedades,df_revision)[0],ingresos(path_novedades,df_revision)[1]]).reset_index(drop=True)

    
    df_revision=crea_archivo_pruebas(path_archivo_base,distribucion_salario(df_revision))
    

    return df_revision

def generarRuta(fecha_corte,rutaInput):
    

    raiz = rutaInput
    rutaUser = raiz[0:raiz[9:len(raiz)].index("/")+9]

    #ruta = rutaUser+'/MVM Ingenieria de Software/Unidad de Sostenibilidad y Crecimiento - Documentos/Confidencial/Privada/Gestión Financiera/Balanc/2022/Control Interno/Nomina/'+ fecha_corte[:10].split("-")[1]+'/REVISIONES/Revision pago de nomina '+fecha_corte.replace("-","")+'.xlsx'
    #ruta = rutaUser+'/MVM Ingenieria de Software/Unidad de Sostenibilidad y Crecimiento - Balanc/'+ fecha_corte[:4]+'/Control Interno/Nomina/'+ fecha_corte[:10].split("-")[1]+'/REVISIONES/Revisión preliminar de nómina '+fecha_corte.replace("-","")+'.xlsx'
    ruta = '.\\resultados\\Revision prueba-2.xlsx'
    return ruta

def calculo_conceptos(path_novedades,path_facturacion,df_pruebas):

    
    

    """
    Esta función calcula sobre el archivo base los diferentes conceptos de nómina por cada empleado. \n
    Adicionalmente, crea una columna de revisión por cada concepto con el objetivo de encontrar errores.

    """

    #Leer archivo 14 para extraer información necesaria:
    xls = pd.ExcelFile(path_novedades)

    # archivo disponibilidades:
    df14_1=xls.parse(2,header=3).iloc[:,[0,1,5,6,7,8]]
    df14_1.rename(columns={df14_1.columns[0]:"CC",df14_1.columns[1]:"NOMBRE"},inplace=True)

    ## archivo horas adicionales
    df14_horas=df14_1.iloc[:,0:4].rename(columns={df14_1.columns[2]:"CANTIDADES",df14_1.columns[3]:"VALOR CUOTA"})
    df14_horas["VALOR CUOTA"]=df14_horas["VALOR CUOTA"].astype(float)
    df14_horas=df14_horas.groupby(["CC"])[["VALOR CUOTA"]].sum().reset_index()

    ## archivo dias adicionales
    df14_dias=df14_1.iloc[:,[0,1,4,5]].rename(columns={df14_1.columns[5]:"VALOR CUOTA"})
    df14_dias["VALOR CUOTA"]=df14_dias["VALOR CUOTA"].astype(float)
    df14_dias=df14_dias.groupby(["CC"])[["VALOR CUOTA"]].sum().reset_index()

    # reembolso gastos
    df14_2=xls.parse(3,header=4).iloc[:,[0,1,7,9]]
    df14_2.rename(columns={df14_2.columns[0]:"CC",df14_2.columns[1]:"NOMBRE","VALOR":"VALOR CUOTA"},inplace=True)
    df14_2["VALOR CUOTA"]=df14_2["VALOR CUOTA"].astype(float)
    df14_2=df14_2[df14_2["DESCRIPCIÓN"]=="REEMBOLSO DE GASTOS "]
    df_reembolso=df14_2.iloc[:,[0,3]]


    # Insertamos columnas vacías para agregar información y formulas
    df_pruebas.insert(14, 'SALARIO BASICO ', np.nan)
    df_pruebas.insert(15, 'SALARIO FLI ', np.nan)
    df_pruebas.insert(16, 'VALIDACIÓN SALARIO', np.nan)
    df_pruebas.insert(22, 'DIA DE LA FAMILIA ', np.nan)
    df_pruebas.insert(33, "REVISION DÍAS", np.nan)
    df_pruebas.insert(34, "OBSERVACIONES", np.nan)
    df_pruebas.insert(36, "REVISIÓN 0010", np.nan)
    df_pruebas.insert(38, "REVISIÓN 0015", np.nan)
    df_pruebas.insert(40, "REVISIÓN 0024", np.nan)
    df_pruebas.insert(42, "REVISIÓN 0025", np.nan)
    df_pruebas.insert(46, "VALOR DISP HORAS", list(df_pruebas.merge(df14_horas,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(47, "REVISIÓN 0316", np.nan)
    df_pruebas.insert(49, "VALOR DISP DIAS", list(df_pruebas.merge(df14_dias,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(50, "REVISIÓN 0317", np.nan)


    ## 0340 Leemos el archivo 13, antes de esto debemos abrirlo y cambiar el formato de la hoja de cargue tarjetas
    xls = pd.ExcelFile(path_facturacion)
    df_14=xls.parse(19,header=1)
    df_14=df_14.iloc[:,[1,2,4,7]].dropna()
    df_14.rename(columns={df_14.columns[0]:"CC",df_14.columns[1]:"NOMBRE",df_14.columns[3]:"VALOR CUOTA",
                          df_14.columns[2]:"COD CONCEPTO"},inplace=True)
    df_14["COD CONCEPTO"]=df_14["COD CONCEPTO"].astype("int64")
    df_14=df_14.groupby(["CC","NOMBRE","COD CONCEPTO"])[["VALOR CUOTA"]].sum().reset_index()
    df_14=df_14[df_14["VALOR CUOTA"]>0].reset_index(drop=True)

    #Conservamos solo tarjetas de alimentación y columnas utiles
    df_alim=df_14[df_14["COD CONCEPTO"]==340].iloc[:,[0,3]]

    #Conservamos solo tarjetas de gasolina y columnas utiles
    df_gaso=df_14[df_14["COD CONCEPTO"]==1953].iloc[:,[0,3]]

    #Libranza Davivienda:
    df_19=xls.parse(11).iloc[:,[0,1]].dropna()
    df_19.rename(columns={df_19.columns[0]:"CC",df_19.columns[1]:"VALOR CUOTA"},inplace=True)
    df_19["COD CONCEPTO"]=2704
    df_19["CONCEPTO"]="LIBRANZA DAVIVIENDA"
    df_19=df_19.groupby([df_19.columns[0],df_19.columns[2],df_19.columns[3]])[[df_19.columns[1]]].sum().reset_index()

    #Libranza Bancolombia:
    df_20=xls.parse(13).iloc[:,[0,1]].dropna()
    df_20.rename(columns={df_20.columns[0]:"CC",df_20.columns[1]:"VALOR CUOTA"},inplace=True)
    df_20["COD CONCEPTO"]=2705
    df_20["CONCEPTO"]="LIBRANZA BANCOLOMBIA"
    df_20=df_20.groupby([df_20.columns[0],df_20.columns[2],df_20.columns[3]])[[df_20.columns[1]]].sum().reset_index()

    #Función para hojas del archivo 13:
    def format13(n,codigo,concepto,head):
        df=xls.parse(n,header=head).iloc[:,0:3].dropna()
        df.rename(columns={df.columns[0]:"CC",df.columns[1]:"NOMBRE",df.columns[2]:"VALOR CUOTA"},inplace=True)
        df["COD CONCEPTO"]=codigo
        df["CONCEPTO"]=concepto
        df=df.groupby([df.columns[0],df.columns[1],df.columns[3],df.columns[4]])[[df.columns[2]]].sum().reset_index()
        df
        return df

    #Prestamos comfama:
    df_10=format13(15,2709,"PRESTAMO COMFAMA",0)

    #Servicios COMFAMA:
    df_22=xls.parse(24,header=0).iloc[:,[0,1,5]].dropna()
    df_22.rename(columns={df_22.columns[0]:"CC",df_22.columns[1]:"NOMBRE",df_22.columns[2]:"VALOR CUOTA"},inplace=True)
    df_22["COD CONCEPTO"]=2710
    df_22["CONCEPTO"]="SERVICIOS COMFAMA MATRICULAS"
    df_22=df_22.groupby([df_22.columns[0],df_22.columns[1],df_22.columns[3],df_22.columns[4]])[[df_22.columns[2]]].sum().reset_index()

    #Descuento EMI:
    df_9=xls.parse(10).iloc[:,[0,1,4]].dropna()
    df_9.rename(columns={df_9.columns[0]:"CC",df_9.columns[1]:"NOMBRE",df_9.columns[2]:"VALOR CUOTA"},inplace=True)
    df_9["COD CONCEPTO"]=2712
    df_9["CONCEPTO"]="DESCUENTO EMI"
    df_9=df_9.groupby([df_9.columns[0],df_9.columns[1],df_9.columns[3],df_9.columns[4]])[[df_9.columns[2]]].sum().reset_index()

    #Descuento PREVER:
    df_12=format13(17,2713,"DESCUENTO PREVER",0)

    #POLIZA AUTO
    df_6=format13(5,2714,"POLIZA DE AUTO",0)

    #POLIZA DE VIDA:
    df_7=format13(7,2715,"POLIZA DE VIDA",0)

    #movistar:
    df_11=xls.parse(16,header=1).iloc[:,[0,2,3]].dropna()
    df_11.rename(columns={df_11.columns[0]:"CC",df_11.columns[1]:"NOMBRE",df_11.columns[2]:"VALOR CUOTA"},inplace=True)
    df_11["COD CONCEPTO"]=2724
    df_11["CONCEPTO"]="DESCUENTO MOVISTAR"
    df_11=df_11.groupby([df_11.columns[0],df_11.columns[1],df_11.columns[3],df_11.columns[4]])[[df_11.columns[2]]].sum().reset_index()

    #Poliza SURA PAC
    df_8=format13(8,2737,"POLIZA SURA PAC",0)

    #AHORRO FEMTI:
    df_21=xls.parse(31).iloc[:,[0,1,3]].dropna()
    df_21.rename(columns={df_21.columns[0]:"CC",df_21.columns[1]:"NOMBRE",df_21.columns[2]:"VALOR CUOTA"},inplace=True)
    df_21["COD CONCEPTO"]=2739
    df_21["CONCEPTO"]="AHORRO FEMTI"
    df_21=df_21.groupby([df_21.columns[0],df_21.columns[1],df_21.columns[3],df_21.columns[4]])[[df_21.columns[2]]].sum().reset_index()

    #APORTE VOLUNTARIO:
    df_18=format13(30,3210,"APORTE VOLUNTARIO A PENSION VOLUNTARIA",0)

    #APORTE AFC:
    df_17=format13(29,3213,"APORTE AFC",0)

    #PREPAGADA COLSANITAS
    df_1=format13(0,3602,"DESCUENTO MEDICINA PREPAGADA COLSANITAS",0)

    #PREPAGADA MEDISANITAS
    df_2=format13(1,3604,"DESCUENTO MEDICINA PREPAGADA MEDISANITAS",0)

    #MEDICINA PREPAGADA SURA - GLOBAL
    df_4=format13(3,3607,"MEDICINA PREPAGADA SURA - GLOBAL",0)

    #MEDICINA PREPAGADA SURA - ESPECIAL
    df_3=format13(2,3608,"MEDICINA PREPAGADA SURA - ESPECIAL",0)

    #MEDICINA PREPAGADA SURA - CLASICA
    df_5=format13(4,3609,"MEDICINA PREPAGADA SURA - CLASICA",0)


    #MEDICINA PREPAGADA COOMEVA
    df_16=format13(22,3611,"DESCUENTO MEDICINA PREPAGADA COOMEVA",0)

    # Insert an empty column to write the formulas
    df_pruebas.insert(52, "VALOR ALIMENTACIÓN", list(df_pruebas.merge(df_alim,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(53, "REVISIÓN 0340", np.nan)
    df_pruebas.insert(55, "REVISIÓN 0740", np.nan)
    df_pruebas.insert(57, "REVISIÓN 0791", np.nan)
    df_pruebas.insert(59, "REVISIÓN 0920", np.nan)
    df_pruebas.insert(61, "REVISIÓN 0923", np.nan)
    df_pruebas.insert(63, "VALOR REEMBOLSO GASTOS", list(df_pruebas.merge(df_reembolso,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(64, "REVISIÓN 1200", np.nan)
    df_pruebas.insert(66, "REVISIÓN 1600", np.nan)
    df_pruebas.insert(68, "REVISIÓN 1670", np.nan)
    df_pruebas.insert(70, "REVISIÓN 1671", np.nan)
    df_pruebas.insert(72, "REVISIÓN 1673", np.nan)
    df_pruebas.insert(74, "REVISIÓN 1730", np.nan)
    df_pruebas.insert(77, "REVISIÓN 1950", np.nan)
    df_pruebas.insert(79, "VALOR GASOLINA", list(df_pruebas.merge(df_gaso,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(80, "REVISIÓN 1953", np.nan)
    df_pruebas.insert(82, "REVISIÓN 1954", np.nan)
    df_pruebas.insert(84, "REVISIÓN 1956", np.nan)
    df_pruebas.insert(86, "REVISIÓN 1957", np.nan)
    df_pruebas.insert(88, "REVISIÓN 1970", np.nan)
    df_pruebas.insert(90, "REVISIÓN 1971", np.nan)
    df_pruebas.insert(93, "REVISIÓN 2500", np.nan)
    df_pruebas.insert(95, "REVISIÓN 2510", np.nan)
    df_pruebas.insert(98, "REVISIÓN 2520", np.nan)
    df_pruebas.insert(100, "VALOR LIBRANZA DAVIVIENDA", list(df_pruebas.merge(df_19,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(101, "REVISIÓN 2704", np.nan)
    df_pruebas.insert(103, "VALOR LIBRANZA BANCOLOMBIA", list(df_pruebas.merge(df_20,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(104, "REVISIÓN 2705", np.nan)
    df_pruebas.insert(106, "VALOR PRESTAMOS COMFAMA", list(df_pruebas.merge(df_10,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(107, "REVISIÓN 2709", np.nan)
    df_pruebas.insert(109, "VALOR SERVICIOS COMFAMA", list(df_pruebas.merge(df_22,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(110, "REVISIÓN 2710", np.nan)
    df_pruebas.insert(112, "VALOR DESCUENTO EMI", list(df_pruebas.merge(df_9,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(113, "REVISIÓN 2712", np.nan)
    df_pruebas.insert(115, "VALOR DESCUENTO PREVER", list(df_pruebas.merge(df_12,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(116, "REVISIÓN 2713", np.nan)
    df_pruebas.insert(118, "VALOR POLIZA AUTO", list(df_pruebas.merge(df_6,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(119, "REVISIÓN 2714", np.nan)
    df_pruebas.insert(121, "VALOR POLIZA VIDA", list(df_pruebas.merge(df_7,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(122, "REVISIÓN 2715", np.nan)
    df_pruebas.insert(124, "VALOR DESCUENTO MOVISTAR", list(df_pruebas.merge(df_11,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(125, "REVISIÓN 2724", np.nan)
    df_pruebas.insert(127, "VALOR POLIZA SURA PAC", list(df_pruebas.merge(df_8,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(128, "REVISIÓN 2737", np.nan)
    df_pruebas.insert(130, "VALOR AHORRO FEMTI", list(df_pruebas.merge(df_21,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(131, "REVISIÓN 2739", np.nan)
    df_pruebas.insert(133, "VALOR APORTE VOLUNTARIO", list(df_pruebas.merge(df_18,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(134, "REVISIÓN 3210", np.nan)
    df_pruebas.insert(136, "VALOR APORTE AFC", list(df_pruebas.merge(df_17,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(137, "REVISIÓN 3213", np.nan)
    df_pruebas.insert(139, "REVISIÓN 3222", np.nan)
    df_pruebas.insert(141, "VALOR PREPAGADA COLSANITAS", list(df_pruebas.merge(df_1,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(142, "REVISIÓN 3602", np.nan)
    df_pruebas.insert(144, "VALOR PREPAGADA MEDISANITAS", list(df_pruebas.merge(df_2,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(145, "REVISIÓN 3604", np.nan)
    df_pruebas.insert(147, "VALOR PREPAGADA SURA GLOBAL", list(df_pruebas.merge(df_4,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(148, "REVISIÓN 3607", np.nan)
    df_pruebas.insert(150, "VALOR PREPAGADA SURA ESPECIAL", list(df_pruebas.merge(df_3,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(151, "REVISIÓN 3608", np.nan)
    df_pruebas.insert(153, "VALOR PREPAGADA SURA CLASICA", list(df_pruebas.merge(df_5,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(154, "REVISIÓN 3609", np.nan)
    df_pruebas.insert(156, "VALOR PREPAGADA COOMEVA", list(df_pruebas.merge(df_16,on="CC",how="left")["VALOR CUOTA"].fillna(0)))
    df_pruebas.insert(157, "REVISIÓN 3611", np.nan)
    df_pruebas.insert(159, "REVISIÓN 3900", np.nan)
    df_pruebas.insert(161, "REVISIÓN 3901", np.nan)
    df_pruebas.insert(163, "REVISIÓN 3903", np.nan)
    df_pruebas.insert(166, "REVISIÓN 3920", np.nan)
    
    global merged_data,df_IBC,dftest,dfinicial,df_merged
    
    dfinicial = df_pruebas
    mes = fecha.split('-')[1]
    df_IBC = agregar_columna_IBC(mes)
    merged_data = pd.merge(df_IBC, df_pruebas, on='CC')
    # df_pruebas = df_pruebas.join(merged_data['IBC'])
    dftest = df_pruebas
    
    # ----INICIO CODIGO DE MERGE FUNCIONANDO
    # Unir los dataframes usando el método merge y especificar que es unir por la columna 'CC'
    df_merged = pd.merge(dfinicial, df_IBC[['CC', 'IBC']], on='CC', how='left')

    # Reemplazar los valores NaN, es decir los registros que no están en df1, con el valor 99
    df_merged['IBC'] = df_merged['IBC'].fillna(0)
    
    #Eliminar comas y convertir a np.float
    df_merged['IBC'] = df_merged['IBC'].replace(',', '', regex=True).astype(np.float64)
    # ----FIN CODIGO DE MERGE FUNCIONANDO
    
    # Start the xlsxwriter
    #writer = pd.ExcelWriter('resultados/NOMINA/2022/Formato revisión nómina'+mes+'3.xlsx', engine='xlsxwriter')
    #writer = pd.ExcelWriter(ruta_guardar+'/Formato revisión preliminar de nómina.xlsx', engine='xlsxwriter')

    rutaGuardarArchivos = generarRuta(fecha,ruta_nomina_horizontal)

    writer = pd.ExcelWriter(rutaGuardarArchivos, engine="xlsxwriter")

    # df_pruebas.to_excel(writer, sheet_name='Hoja1', index=False)
    df_merged.to_excel(writer, sheet_name='Hoja1', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Hoja1']


    # Create a for loop to start writing the formulas to each row
    global col_IBC
    col_IBC = num_a_col_excel(df_merged.columns.get_loc('IBC')+1)

    #Formulas:
        

    #Salario basico
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=N{row}*L{row}'
        worksheet.write_formula(f"O{row}", formula)

    #Salario fli
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=N{row}*M{row}'
        worksheet.write_formula(f"P{row}", formula)

    #Validación Salario
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(O{row}+P{row})-N{row}'
        worksheet.write_formula(f"Q{row}", formula)

    #agregamos día de la familia:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AN{row}=0,0,1)'
        worksheet.write_formula(f"W{row}", formula)

    #agregamos REVISIÓN DÍAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=SUM(R{row}:AF{row})-T{row}'
        worksheet.write_formula(f"AH{row}", formula)

    #agregamos REVISIÓN 0010
    for row in range(2,df_pruebas.shape[0]+2): 
        formula = f'=IF(F{row}="LEY 50",((O{row}/30)*R{row})-AJ{row},0)'
        worksheet.write_formula(f"AK{row}", formula)

    #agregamos REVISIÓN 0015
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",((O{row}/30)*R{row})-AL{row},0)'
        worksheet.write_formula(f"AM{row}", formula)

    #agregamos REVISIÓN 0024
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*W{row})-AN{row}'
        worksheet.write_formula(f"AO{row}", formula)

    # Agregamos REVISIÓN 0025
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="APRENDIZ",((O{row}/30)*R{row})-AP{row},0)'
        worksheet.write_formula(f"AQ{row}", formula)

    # REVISIÓN 0316:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AU{row}-AT{row}'
        worksheet.write_formula(f"AV{row}", formula)

    #0317-ATENCION DISPONIBILIDAD DIAS:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AX{row}-AW{row}'
        worksheet.write_formula(f"AY{row}", formula)

    # 0340-ATARJETA ALIMENTACIÓN:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BA{row}-AZ{row}'
        worksheet.write_formula(f"BB{row}", formula)

    # 0740-AUXILIO EMPRESA INCAPACIDAD:
    # for row in range(2,df_pruebas.shape[0]+2):
    #     formula = f'=IF(AND(S{row}<=2,F{row}<>"INTEGRAL"),((O{row}/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}<>"INTEGRAL"),((O{row}/30)*2)-BC{row},IF(AND(S{row}<=2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*2)-BC{row},0))))'
    #     worksheet.write_formula(f"BD{row}", formula)


    # 0740-AUXILIO EMPRESA INCAPACIDAD: EDITADO LUIS PEREZ
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF({col_IBC}{row}>0,IF(AND(S{row}<=2,F{row}<>"INTEGRAL"),(({col_IBC}{row}/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}<>"INTEGRAL"),(({col_IBC}{row}/30)*2)-BC{row},IF(AND(S{row}<=2,F{row}="INTEGRAL"),((({col_IBC}{row}*70%)/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}="INTEGRAL"),((({col_IBC}{row}*70%)/30)*2)-BC{row},0)))),IF(AND(S{row}<=2,F{row}<>"INTEGRAL"),((O{row}/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}<>"INTEGRAL"),((O{row}/30)*2)-BC{row},IF(AND(S{row}<=2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*2)-BC{row},0)))))'
        worksheet.write_formula(f"BD{row}", formula)




    # 0791 AUXILIO DE CONECTIVIDAD: Solo a quienes ganen menos de 2 smmlv
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AND(O{row}<=2000000,BE{row}<>0),True,False)'
        worksheet.write_formula(f"BF{row}", formula)

    # 0920 AVACACIONES EN DINERO:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*T{row})-BG{row}'
        worksheet.write_formula(f"BH{row}", formula)

    # 0923-AUXILIO VACACIONES RFI: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((P{row}/30)*T{row})-BI{row}'
        worksheet.write_formula(f"BJ{row}", formula)

    # 1200-REEMBOLSO DE GASTOS: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BL{row}-BK{row}'
        worksheet.write_formula(f"BM{row}", formula)

    # 1600-ENFERMEDAD GENERAL: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AND(S{row}>=0,S{row}<=2),0,((N{row}/30)*S{row}-2)-BN{row})'
        worksheet.write_formula(f"BO{row}", formula)

    #1670-LICENCIA REMUNERADA  
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*AA{row})-BP{row}'
        worksheet.write_formula(f"BQ{row}", formula)

    # 1671-CALAMIDAD DOMESTICA    
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*Z{row})-BR{row}'
        worksheet.write_formula(f"BS{row}", formula)

    # 1673-LICENCIA POR LUTO  
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*Z{row})-BT{row}'
        worksheet.write_formula(f"BU{row}", formula)

    # 1730-VACACIONES DISFRUTADAS 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*AF{row})-BV{row}'
        worksheet.write_formula(f"BW{row}", formula)

    # 1950-VAPORTE VOLUNTARIO INST.
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}*4%)-BY{row}'
        worksheet.write_formula(f"BZ{row}", formula)

    #1953- TARJETA GASOLINA RFI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CB{row}-CA{row}'
        worksheet.write_formula(f"CC{row}", formula)

    #1954- BENEFICIO PLAN
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CE{row}", formula)

    #1956- BENEFICIO PLAN 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CG{row}", formula)

    #1957- FLEX BEN APOR 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CI{row}", formula)

    #1970-APORTE INSTITUCIONAL 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}/30*R{row})*0.075-CJ{row}' 
        worksheet.write_formula(f"CK{row}", formula)

    #1971-APORTE VOLUNTARIO PLUS 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}/30*R{row})*0.1766-CL{row}' 
        worksheet.write_formula(f"CM{row}", formula)

    #2500-APORTE SALUD EGM 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.04,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.04)-CO{row}' 
        worksheet.write_formula(f"CP{row}", formula)

    #2510-FONDO DE SOLIDARIDAD
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AJ{row}>4000000,IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.01,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.01),0)-CQ{row}'
        worksheet.write_formula(f"CR{row}", formula)

    #2520-APORTES PENSION IVM
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.04,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.04)-CT{row}'
        worksheet.write_formula(f"CU{row}", formula)

    #2704-LIBRANZA DAVIVIENDA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CW{row}-CV{row}'
        worksheet.write_formula(f"CX{row}", formula)

    #2705-LIBRANZA BANCOLOMBIA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CZ{row}-CY{row}'
        worksheet.write_formula(f"DA{row}", formula)

    #2709-PRESTAMO COMFAMA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DC{row}-DB{row}'
        worksheet.write_formula(f"DD{row}", formula)

    #2710-SERVICIOS COMFAMA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DF{row}-DE{row}'
        worksheet.write_formula(f"DG{row}", formula)

    #2712-DESCUENTO EMI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DI{row}-DH{row}'
        worksheet.write_formula(f"DJ{row}", formula)

    #2713-DESCUENTO PREVER
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DL{row}-DK{row}'
        worksheet.write_formula(f"DM{row}", formula)

    #2714-POLIZA DE AUTO
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DO{row}-DN{row}'
        worksheet.write_formula(f"DP{row}", formula)

    #2715-POLIZA DE VIDA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DR{row}-DQ{row}'
        worksheet.write_formula(f"DS{row}", formula)

    #2724-DESCUENTO MOVISTAR
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DU{row}-DT{row}'
        worksheet.write_formula(f"DV{row}", formula)

    #2737-POLIZA SURA PAC
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DX{row}-DW{row}'
        worksheet.write_formula(f"DY{row}", formula)

    #2739-AHORRO FEMTI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EA{row}-DZ{row}'
        worksheet.write_formula(f"EB{row}", formula)

    #3210-APORTE VOLUNTARIO A
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=ED{row}-EC{row}'
        worksheet.write_formula(f"EE{row}", formula)

    #3213-APORTE AFC
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EG{row}-EF{row}'
        worksheet.write_formula(f"EH{row}", formula)

    #3222-APORTE VOLUNTARIO PLUS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CL{row}-EI{row}'
        worksheet.write_formula(f"EJ{row}", formula)

    #3602-DESCUENTO MEDICINA COLSANITAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EL{row}-EK{row}'
        worksheet.write_formula(f"EM{row}", formula)

    #3604-DESCUENTO MEDICINA MEDISANITAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EO{row}-EN{row}'
        worksheet.write_formula(f"EP{row}", formula)

    #3607-MEDICINA PREPAGADA SURA - GLOBAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=ER{row}-EQ{row}'
        worksheet.write_formula(f"ES{row}", formula)

    #3608-MEDICINA PREPAGADA SURA - ESPECIAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EU{row}-ET{row}'
        worksheet.write_formula(f"EV{row}", formula)

    #3609-MEDICINA PREPAGADA SURA - CLASICA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EX{row}-EW{row}'
        worksheet.write_formula(f"EY{row}", formula)

    #3611-MEDICINA PREPAGADA COOMEVA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=FA{row}-EZ{row}'
        worksheet.write_formula(f"FB{row}", formula)

    #3900-APORTE VOLUNTARIO INST. 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BY{row}-FC{row}'
        worksheet.write_formula(f"FD{row}", formula)

    #3901-TARJETA ALIMENTACIÓN 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AZ{row}-FE{row}'
        worksheet.write_formula(f"FF{row}", formula)

    #3903-TARJETA GASOLINA RFI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CA{row}-FG{row}'
        worksheet.write_formula(f"FH{row}", formula)

    #3920-APORTE INSTITUCIONAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CJ{row}-FJ{row}'
        worksheet.write_formula(f"FK{row}", formula)


    writer.save()
    writer.close()
    
    #calcular_0740()
    
#----EDITADO POR LUIS ALEJANDRO PÉREZ ABRIL-----------------------------------#
#----REVISIÓN DIFERENCIAS INCAPACIDADES---------------------------------------#

#Este método evaluará de nuevo el concepto 0740, atenderá los casos que hayan 
#tenido aumento de salario 
def leer_RevPreliminar():
    
    
    xls = pd.ExcelFile('.\\resultados\Revision prueba-2.xlsm')

    df14_16=xls.parse(0,header=0)
    
    
    return df14_16

def leer_RevPreliminarxlsx():
    
    
    xls = pd.ExcelFile('.\\resultados\Revision prueba-2.xlsx')

    df14_16=xls.parse(0,header=0)
    
    
    return df14_16

def leer_NovedadesNomina():
    
    
    xls = pd.ExcelFile('.\\resultados\\14.NOVEDADES DE NOMINA 20230201.xlsm')
    
    df = pd.read_excel(xls, sheet_name='VARIACIÓN SALARIO', header=2)

    
    
    return df

def agregar_formulas(df_pruebas,worksheet,writer):
    
    #Formulas:

    #Salario basico
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=N{row}*L{row}'
        worksheet.write_formula(f"O{row}", formula)

    #Salario fli
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=N{row}*M{row}'
        worksheet.write_formula(f"P{row}", formula)

    #Validación Salario
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(O{row}+P{row})-N{row}'
        worksheet.write_formula(f"Q{row}", formula)

    #agregamos día de la familia:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AN{row}=0,0,1)'
        worksheet.write_formula(f"W{row}", formula)

    #agregamos REVISIÓN DÍAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=SUM(R{row}:AF{row})-T{row}'
        worksheet.write_formula(f"AH{row}", formula)

    #agregamos REVISIÓN 0010
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="LEY 50",((O{row}/30)*R{row})-AJ{row},0)'
        worksheet.write_formula(f"AK{row}", formula)

    #agregamos REVISIÓN 0015
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",((O{row}/30)*R{row})-AL{row},0)'
        worksheet.write_formula(f"AM{row}", formula)

    #agregamos REVISIÓN 0024
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*W{row})-AN{row}'
        worksheet.write_formula(f"AO{row}", formula)

    # Agregamos REVISIÓN 0025
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="APRENDIZ",((O{row}/30)*R{row})-AP{row},0)'
        worksheet.write_formula(f"AQ{row}", formula)

    # REVISIÓN 0316:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AU{row}-AT{row}'
        worksheet.write_formula(f"AV{row}", formula)

    #0317-ATENCION DISPONIBILIDAD DIAS:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AX{row}-AW{row}'
        worksheet.write_formula(f"AY{row}", formula)

    # 0340-ATARJETA ALIMENTACIÓN:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BA{row}-AZ{row}'
        worksheet.write_formula(f"BB{row}", formula)

    # 0740-AUXILIO EMPRESA INCAPACIDAD:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AND(S{row}<=2,F{row}<>"INTEGRAL"),((O{row}/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}<>"INTEGRAL"),((O{row}/30)*2)-BC{row},IF(AND(S{row}<=2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*S{row})-BC{row},IF(AND(S{row}>2,F{row}="INTEGRAL"),(((O{row}*70%)/30)*2)-BC{row},0))))'
        worksheet.write_formula(f"BD{row}", formula)

    # 0791 AUXILIO DE CONECTIVIDAD: Solo a quienes ganen menos de 2 smmlv
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AND(O{row}<=2000000,BE{row}<>0),True,False)'
        worksheet.write_formula(f"BF{row}", formula)

    # 0920 AVACACIONES EN DINERO:
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*T{row})-BG{row}'
        worksheet.write_formula(f"BH{row}", formula)

    # 0923-AUXILIO VACACIONES RFI: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((P{row}/30)*T{row})-BI{row}'
        worksheet.write_formula(f"BJ{row}", formula)

    # 1200-REEMBOLSO DE GASTOS: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BL{row}-BK{row}'
        worksheet.write_formula(f"BM{row}", formula)

    # 1600-ENFERMEDAD GENERAL: 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AND(S{row}>=0,S{row}<=2),0,((N{row}/30)*S{row}-2)-BN{row})'
        worksheet.write_formula(f"BO{row}", formula)

    #1670-LICENCIA REMUNERADA  
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*AA{row})-BP{row}'
        worksheet.write_formula(f"BQ{row}", formula)

    # 1671-CALAMIDAD DOMESTICA    
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*Z{row})-BR{row}'
        worksheet.write_formula(f"BS{row}", formula)

    # 1673-LICENCIA POR LUTO  
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*Z{row})-BT{row}'
        worksheet.write_formula(f"BU{row}", formula)

    # 1730-VACACIONES DISFRUTADAS 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=((O{row}/30)*AF{row})-BV{row}'
        worksheet.write_formula(f"BW{row}", formula)

    # 1950-VAPORTE VOLUNTARIO INST.
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}*4%)-BY{row}'
        worksheet.write_formula(f"BZ{row}", formula)

    #1953- TARJETA GASOLINA RFI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CB{row}-CA{row}'
        worksheet.write_formula(f"CC{row}", formula)

    #1954- BENEFICIO PLAN
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CE{row}", formula)

    #1956- BENEFICIO PLAN 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CG{row}", formula)

    #1957- FLEX BEN APOR 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=P{row}-AS{row}-AZ{row}-BY{row}-CA{row}-CD{row}-CF{row}-CH{row}'
        worksheet.write_formula(f"CI{row}", formula)

    #1970-APORTE INSTITUCIONAL 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}/30*R{row})*0.075-CJ{row}' 
        worksheet.write_formula(f"CK{row}", formula)

    #1971-APORTE VOLUNTARIO PLUS 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=(P{row}/30*R{row})*0.1766-CL{row}' 
        worksheet.write_formula(f"CM{row}", formula)

    #2500-APORTE SALUD EGM 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.04,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.04)-CO{row}' 
        worksheet.write_formula(f"CP{row}", formula)

    #2510-FONDO DE SOLIDARIDAD
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(AJ{row}>4000000,IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.01,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.01),0)-CQ{row}'
        worksheet.write_formula(f"CR{row}", formula)

    #2520-APORTES PENSION IVM
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=IF(F{row}="INTEGRAL",(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.7*0.04,(AJ{row}+AL{row}+AN{row}+AT{row}+AW{row}+BC{row}+BN{row}+BP{row}+BR{row}+BT{row}+BV{row})*0.04)-CT{row}'
        worksheet.write_formula(f"CU{row}", formula)

    #2704-LIBRANZA DAVIVIENDA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CW{row}-CV{row}'
        worksheet.write_formula(f"CX{row}", formula)

    #2705-LIBRANZA BANCOLOMBIA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CZ{row}-CY{row}'
        worksheet.write_formula(f"DA{row}", formula)

    #2709-PRESTAMO COMFAMA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DC{row}-DB{row}'
        worksheet.write_formula(f"DD{row}", formula)

    #2710-SERVICIOS COMFAMA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DF{row}-DE{row}'
        worksheet.write_formula(f"DG{row}", formula)

    #2712-DESCUENTO EMI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DI{row}-DH{row}'
        worksheet.write_formula(f"DJ{row}", formula)

    #2713-DESCUENTO PREVER
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DL{row}-DK{row}'
        worksheet.write_formula(f"DM{row}", formula)

    #2714-POLIZA DE AUTO
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DO{row}-DN{row}'
        worksheet.write_formula(f"DP{row}", formula)

    #2715-POLIZA DE VIDA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DR{row}-DQ{row}'
        worksheet.write_formula(f"DS{row}", formula)

    #2724-DESCUENTO MOVISTAR
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DU{row}-DT{row}'
        worksheet.write_formula(f"DV{row}", formula)

    #2737-POLIZA SURA PAC
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=DX{row}-DW{row}'
        worksheet.write_formula(f"DY{row}", formula)

    #2739-AHORRO FEMTI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EA{row}-DZ{row}'
        worksheet.write_formula(f"EB{row}", formula)

    #3210-APORTE VOLUNTARIO A
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=ED{row}-EC{row}'
        worksheet.write_formula(f"EE{row}", formula)

    #3213-APORTE AFC
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EG{row}-EF{row}'
        worksheet.write_formula(f"EH{row}", formula)

    #3222-APORTE VOLUNTARIO PLUS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CL{row}-EI{row}'
        worksheet.write_formula(f"EJ{row}", formula)

    #3602-DESCUENTO MEDICINA COLSANITAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EL{row}-EK{row}'
        worksheet.write_formula(f"EM{row}", formula)

    #3604-DESCUENTO MEDICINA MEDISANITAS
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EO{row}-EN{row}'
        worksheet.write_formula(f"EP{row}", formula)

    #3607-MEDICINA PREPAGADA SURA - GLOBAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=ER{row}-EQ{row}'
        worksheet.write_formula(f"ES{row}", formula)

    #3608-MEDICINA PREPAGADA SURA - ESPECIAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EU{row}-ET{row}'
        worksheet.write_formula(f"EV{row}", formula)

    #3609-MEDICINA PREPAGADA SURA - CLASICA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=EX{row}-EW{row}'
        worksheet.write_formula(f"EY{row}", formula)

    #3611-MEDICINA PREPAGADA COOMEVA
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=FA{row}-EZ{row}'
        worksheet.write_formula(f"FB{row}", formula)

    #3900-APORTE VOLUNTARIO INST. 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=BY{row}-FC{row}'
        worksheet.write_formula(f"FD{row}", formula)

    #3901-TARJETA ALIMENTACIÓN 
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=AZ{row}-FE{row}'
        worksheet.write_formula(f"FF{row}", formula)

    #3903-TARJETA GASOLINA RFI
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CA{row}-FG{row}'
        worksheet.write_formula(f"FH{row}", formula)

    #3920-APORTE INSTITUCIONAL
    for row in range(2,df_pruebas.shape[0]+2):
        formula = f'=CJ{row}-FJ{row}'
        worksheet.write_formula(f"FK{row}", formula)
        
    writer.save()
    writer.close()

def agregar_columna_IBC(mes):
    
    # Restar 1 dia al mes actual, para tomar ibc del mes anterior
    mesInt = int(mes)
    mesIbc = mesInt -1

    if mesIbc == 0:
        mesIbc = '12'
    else:
        mesIbc = str(mesIbc).zfill(2)

    # Leer archivo de ibc
    df = pd.read_excel(".\\resultados\IBC.xls", sheet_name="ACUMU_INFOR_XEMPLO", header=7)

    # Eliminar filas y columnas basura
    column_names = df.columns
    columns_to_drop = [index for index, name in enumerate(column_names) if 'Unnamed' in name]
    
    global dataframe
    
    dataframe = df.drop(df.columns[columns_to_drop], axis=1)
    dataframe = dataframe.dropna(how='all')


    # Econtrar cual es el mes de interes para tomar el ibc, toma el mes actual - 1
    columnasDf = dataframe.columns
    posicion = None

    for i, valor in enumerate(columnasDf):
        if valor.find("-"+mesIbc+"-") != -1:
            posicion = i
            break

    if posicion is not None:
        columnaIBC = columnasDf[posicion]
    else:
        columnaIBC = None
    
    global df_ibc
    df_ibc = pd.DataFrame(columns=['CODIGO', 'IBC'])

    # Recorrer las filas para organizar el dataframe, separar por código e IBC
    valueCodigo = ''
    count = 0
    # recorrer las filas del dataframe original
    for index, row in dataframe.iterrows():
        count = count +1
        # verificar si la longitud de CODIGO es mayor a 4
        if len(row['CODIGO']) > 4:
            nueva_fila = {'CODIGO': row['CODIGO'], 'IBC': '0'}
            valueCodigo = row['CODIGO']
        else:
            # verificar si el valor de CODIGO es igual a '9500'
            if row['CODIGO'] == '9500':
                nueva_fila = {'CODIGO': valueCodigo, 'IBC': row[columnaIBC]}
                nueva_fila_df = pd.DataFrame.from_dict(nueva_fila, orient="index").T
                df_ibc = pd.concat([df_ibc, nueva_fila_df], ignore_index=True)
    # Quitar el digito de verificacion de cc 

    def limpiar_codigo(codigo):
        # Separamos el código por las guiones -
        codigo_sep = codigo.split("-")
        # Si la cantidad de guiones es mayor a 2, solo nos quedamos con el primero y el último
        if len(codigo_sep) > 2:
            codigo_sep = [codigo_sep[0], codigo_sep[-1]]
        # Unimos nuevamente el código con un guion -
        codigo_limpio = "-".join(codigo_sep)
        return codigo_limpio

    # Aplicamos la función a la columna CODIGO
    df_ibc['CODIGO'] = df_ibc['CODIGO'].apply(limpiar_codigo)
    # Dividir la columna 'codigo' en dos utilizando el carácter '-'
    codigo_dividido = df_ibc['CODIGO'].str.split('-', expand=True)
    # Asignar la primera columna al nuevo dataframe 'df_nuevo' como la columna 'CC'
    df_ibc = df_ibc.assign(CC=codigo_dividido[0])
    # Asignar la segunda columna al dataframe 'df_nuevo' como la columna 'NOMBRE'
    df_ibc['NOMBRE'] = codigo_dividido[1]
    # Quitar los espacios en blanco al principio y al final de la columna 'NOMBRE'
    df_ibc['NOMBRE'] = df_ibc['NOMBRE'].str.strip()
    df_ibc = df_ibc[['CC','NOMBRE','IBC']]
    df_ibc['CC'] = df_ibc['CC'].apply(int)
    return df_ibc

    
def calcular_0740():
    
    def calculo_fila(fila):
        salario_total = fila["SALARIO BASICO "]
        dias_incapacidad = fila["DIAS INCAPACIDAD"]
        if dias_incapacidad > 2:
          result = salario_total / 30 * 2
        else:
          result = salario_total / 30 * dias_incapacidad
        return result
    
    df_Novedades_Nomina = leer_NovedadesNomina()
    df_RevPrel = leer_RevPreliminar()
    
    global resultados,df_filtrado
    df_filtrado = df_RevPrel.loc[abs(df_RevPrel['REVISIÓN 0740']) > 1000.0]
    resultados = df_filtrado.apply(calculo_fila, axis=1)
    df_filtrado["0740-AJUSTADO"] = resultados

    df_RevPrelxlsx = leer_RevPreliminar()
    cc_en_df4 = set(df_filtrado['CC'])
    # Ahora, para cada cc en dfrevpreli que esté en df4, reemplazaremos el salario en dfrevpreli con el de df4
    for i, row in df_RevPrelxlsx.iterrows():
        if row['CC'] in cc_en_df4:
            salario_actual = df_filtrado.loc[df_filtrado['CC'] == row['CC'], '0740-AJUSTADO'].iloc[0]
            df_RevPrelxlsx.at[i, '0740-AUXILIO EMPRESA '] = salario_actual
        
    #--------------

    writer = pd.ExcelWriter('.\\resultados\\Revision prueba-2DEMO.xlsx', engine="xlsxwriter")
    df_RevPrelxlsx.to_excel(writer, sheet_name='Hoja1', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Hoja1']
    agregar_formulas(df_RevPrelxlsx,worksheet,writer)
    #--------------
    #writer = pd.ExcelWriter('output.xlsx')
    
    # Write the DataFrame to the Excel file
    #df_RevPrelxlsx.to_excel(writer)

    # Save the Excel file
    #writer.save()
  




#----Interfaz-----------------------------------------------------------------#
def interfaz():
    global fecha
    #global ruta_guardar
    #sg.theme("DarkBlue3")
    sg.set_options(font=("Microsoft JhengHei", 13),background_color=('#ffffff'))

    # Add your new theme colors and settings
    sg.LOOK_AND_FEEL_TABLE['MyCreatedTheme'] = {'BACKGROUND': '#ffffff',
                                            'TEXT': '#929292',
                                            'INPUT': '#ffffff',
                                            'TEXT_INPUT': '#929292',
                                            'SCROLL': '#99CC99',
                                            'BUTTON': ('#fff', '#4C9C2E'),
                                            'PROGRESS': ('#D1826B', '#CC8019'),
                                            'BORDER': 1, 'SLIDER_DEPTH': 0, 
    'PROGRESS_DEPTH': 0, }

    # Switch to use your newly created theme
    sg.theme('MyCreatedTheme')

    layout = [

    
        [   
            sg.Image('./ICONO M.png', size=(64,90),background_color=('#ffffff')),
            #sg.Text("Revisión preliminar de Nómina",background_color='#fff',text_color='#4C9C2E',font=('bold',37)),#\n\nVoy a ayudarte a crear el archivo para la revisión preliminar de la nómina. Para esto necesito que me indiques una fecha y\nque se adjunten los archivos que se solicitan a continuación (los archivos deben ser con corte al mes que se quiere revisar):"),
            sg.Text("Revisión preliminar de Nómina",text_color='#929292',background_color='#fff',font=('bold',37)),#\n\nVoy a ayudarte a crear el archivo para la revisión preliminar de la nómina. Para esto necesito que me indiques una fecha y\nque se adjunten los archivos que se solicitan a continuación (los archivos deben ser con corte al mes que se quiere revisar):"),

        ],

        [
            sg.Text("Voy a ayudarte a crear el archivo para la revisión preliminar de la nómina. Para esto se requiere que cargue la información\nque se indica a continuación. Los resultados serán consolidados en un excel que se guardará en la ruta '/MVM Ingenieria\nde Software/Unidad de Sostenibilidad y Crecimiento - Balanc/<año>/Control Interno/Nomina/<mes>/REVISIONES/',\ncon el nombre 'Revisión preliminar de nómina'.",
            background_color='#fff',text_color='#929292',font=("Microsoft JhengHei",13)),
        ],
        [
            sg.Text("Ingreso de información",text_color="#4C9C2E",font=('bold',20),background_color='#fff'),

        ],

        [
            sg.Text("  • Indique útlimo día del mes de corte",size=34,background_color='#fff',text_color='#000000'),
            sg.Input(key="-FECHA-",size=45,font=12,expand_x=True),
            sg.CalendarButton("Calendario",size=9,button_color='#4C9C2E',border_width='1',
            close_when_date_chosen=True,location=(900,100),no_titlebar=False,title="Fecha")

        ],
        [
            sg.Text("  • Nómina horizontal detallada",size=(34),background_color='#fff',text_color='#000000'),
            sg.Input(key='-INPUT_NOMHORI-', enable_events=True,font=(12),size=45,expand_x=True),
            sg.FileBrowse("Buscar",target="-INPUT_NOMHORI-",size=9,file_types=(("Archivos Excel", "*.xlsx"), ("Archivos Excel", "*.xls"),("Archivos Excel", "*.xlsm")),button_color='#4C9C2E'),

        ],
        [
            sg.Text("  • 11. Maestro",size=(34),background_color='#fff',text_color='#000000'),
            sg.Input(key='-INPUT_MAESTRO-', enable_events=True,font=(12),size=45,expand_x=True),
            sg.FileBrowse("Buscar",target="-INPUT_MAESTRO-",size=9,file_types=(("Archivos Excel", "*.xlsx"), ("Archivos Excel", "*.xls"),("Archivos Excel", "*.xlsm")),button_color='#4C9C2E'),

        ],


        [
            sg.Text("  • 13. Facturación nómina",size=(34),background_color='#fff',text_color='#000000'),
            sg.Input(key='-INPUT_FACTURACION-', enable_events=True,font=(12),size=45,expand_x=True),
            sg.FileBrowse("Buscar",target="-INPUT_FACTURACION-",size=9,file_types=(("Archivos Excel", "*.xlsx"), ("Archivos Excel", "*.xls"),("Archivos Excel", "*.xlsm")),button_color='#4C9C2E'),

        ],
        [
            sg.Text("  • 14. Novedades nómina",size=(34),background_color='#fff',text_color='#000000'),
            sg.Input(key='-INPUT_NOVEDADES-', enable_events=True,font=(12),size=45,expand_x=True),
            sg.FileBrowse("Buscar",target="-INPUT_NOVEDADES-",size=9,file_types=(("Archivos Excel", "*.xlsx"), ("Archivos Excel", "*.xls"),("Archivos Excel", "*.xlsm")),button_color='#4C9C2E'),

        ], 
        [
            sg.Text("  • '16. Revision Nomina' del mes anterior",size=(34),background_color='#fff',text_color='#000000'),
            sg.Input(key='-INPUT_ANTERIOR-', enable_events=True,font=(12),size=45,expand_x=True),
            sg.FileBrowse("Buscar",target="-INPUT_ANTERIOR-",size=9,file_types=(("Archivos Excel", "*.xlsx"), ("Archivos Excel", "*.xls"),("Archivos Excel", "*.xlsm")),button_color='#4C9C2E'),

        ],   
   
        [
            sg.Text("Ejecución",text_color="#4C9C2E",font=('bold',20),background_color='#fff'),

        ],

        [
            sg.Text("Después de agregar toda la información solicitada, presione 'Procesar y guardar' →",
            size=65,background_color='#fff',text_color='#929292',font=("Microsoft JhengHei",13)),
            sg.Button("Procesar y guardar",size=18,button_color='#4C9C2E',border_width='2'), 
            sg.CButton("Cerrar",size=9,button_color='#FA4949',border_width='1')

        ],    

    ]



    window = sg.Window('MERCURY - REVISIÓN PRELIMINAR DE NÓMINA', layout)


    df = pd.DataFrame()

    while True:

        event, values = window.read(timeout=100)
        if event == sg.WINDOW_CLOSED:
            break
        elif event is None:
            break
        elif event == 'Procesar y guardar':

            global ruta_nomina_horizontal
            fecha = values['-FECHA-'][:10]        
            ruta_nomina_horizontal = values['-INPUT_NOMHORI-']
            ruta_maestro = values['-INPUT_MAESTRO-']
            ruta_revision_mesAnterior=values["-INPUT_ANTERIOR-"]
            ruta_novedades_nomina=values["-INPUT_NOVEDADES-"]
            ruta_facturacion_nomina=values["-INPUT_FACTURACION-"]
            ruta_archivo_base=".\\archivos\\archivo base de revision.xlsx"

            try:
                df_new=calculo_conceptos(ruta_novedades_nomina,ruta_facturacion_nomina,
                        compila_archivo(
                        ruta_nomina_horizontal,ruta_maestro,ruta_revision_mesAnterior,ruta_novedades_nomina,ruta_archivo_base
                        )
                    )
                
                
                
                
                df=pd.concat([df,df_new])
                
                

                sg.popup("¡El proceso ha finalizado con exito!","Se ha guarado una copia del archivo en la ruta:\n'/MVM Ingenieria de Software/Unidad de Sostenibilidad y Crecimiento - Balanc/"+fecha[:4]+'/Control Interno/Nomina\n/'+ fecha[:10].split("-")[1]+"/REVISIONES/'.\nPor favor cierre la ventana.",title="Resultado",text_color="#4C9C2E",font='bold',background_color='#fff')
                
            except Exception as e:

                sg.popup_error(f'Ha ocurrido el siguiente error: ' , e, 'Por favor verifique que haya cargado todos los archivos con el formato adecuado.',title="Resultado",background_color='#fff',text_color='#FF0505')
                print("Error: ", e)


    window.close()
    
if __name__=="__main__":
    interfaz()