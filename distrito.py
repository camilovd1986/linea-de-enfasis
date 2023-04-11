from colorama import Fore,Back,Style
import pandas as pd
import openpyxl 


factorservi = 0
def init():
    while True:
        try:
            factorservi = float(input(Fore.BLUE + "Ingresar factor de servico entre 1 y 3: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue    
        if(factorservi >= 1 and factorservi <= 3):
            cal(factorservi)        
        else:
            print(Fore.RED + "El factor de servicio no es correcto")
        init()

def cal(parametro):
    while True:
        try:
            caudal = float(input(Fore.BLUE + "Ingresar caudal: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        break
    while True:
        try:        
            tentrada = float(input(Fore.BLUE + "Ingresar Temp entrada: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        
        break 
    while True:
        try:      
            tsalida = float(input(Fore.BLUE + "Ingresar Temp salida: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        
        break    
    consta1 = 1000
    consta2 = 0.0003069

    tamanioDT = round(caudal*(tentrada - tsalida)*parametro*consta1*consta2)

    print(Fore.YELLOW + "\n El distrito mide ", tamanioDT , '\n')
    chillers(tamanioDT)

def chillers(tdt):
    print(Fore.YELLOW + "\n Tamaños de chillers Centrífugos y de Absorción 500TR, 750TR, 1000TR \n")
    print(Fore.BLUE + "Favor indicar cantidad y lea con detenimiento \n")
    print(Fore.MAGENTA + "__________________________________________________________ \n")
    while True:
        try:
            c500 = int(input(Fore.BLUE + "Ingrese cantidad para 500TR centrífugos: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if c500 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break
    while True:
        try:    
            c750 = int(input(Fore.BLUE + "Ingrese cantidad para 750TR centrífugos: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if c750 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break 
    while True:
        try:       
            c1000 = int(input(Fore.BLUE + "Ingrese cantidad para 1000TR centrífugos: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if c1000 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break
    while True:
        try:        
            aa500 = int(input(Fore.BLUE + "Ingrese cantidad para 500TR Absorción: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if aa500 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break
    while True:
        try:        
            a750 = int(input(Fore.BLUE + "Ingrese cantidad para 750TR Absorción: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if a750 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break 
    while True:
        try:       
            a1000 = int(input(Fore.BLUE + "Ingrese cantidad para 1000TR Absorción: "))
        except ValueError:
            print(Fore.RED + "solo digita numeros")
            continue
        if c500 < 0:
            print(Fore.RED + "debes escribir un numero positivo")
            continue
        else:
            break    

    #Operación centrífugos
    totalc= (500*c500)+(750*c750)+(1000*c1000)
    totala= (500*aa500)+(750*a750)+(1000*a1000)
    totales = totala + totalc

    tmax = tdt + (tdt*0.5) #Se comprueba el tamaño maximo de TR
    if totales<=tdt:
        print(Fore.RED + "\n Las tecnologías seleccionadas no suministran el tamaño del DT \n")
        print(Fore.MAGENTA + "__________________________________________________________ \n")
        chillers(tdt)
    elif totales >= tmax:
        print(Fore.RED + "\n Las tecnologías seleccionadas superan el tope del DT")
        print(Fore.MAGENTA + "__________________________________________________________ \n")
        chillers(tdt)
    else:
        centrifugos(totalc)
        absorcion(totala)

def centrifugos(parametro1):
    
    rp=parametro1*0.3190995427365	
    g=(parametro1*511.13199046407)/1000	
    c=(parametro1*0.0035174111853)*(1925000/0.88)	
    o=c*0.03	
    	
    capex=parametro1*0.0035174111853	
    ft=capex*1000000	
    e=capex*1700000	
    b=capex*2000000
    # Creamos la tabla centrífugos
    centri = {'Energia': ['Red Publica', 'Microturbina Gas', 'Solar Foto Voltaica', 'Energia Eolica','Energia Biomasa','TR de los chillers centrifugos es:'],
              'Emisiones CO2(tco2 al mes)':[e,rp,b,ft,c,0.0003069],
                'CAPEX(dolares megavatios)':[g,o,b,ft,capex,""],
            'opex(do-año)': [ft,rp,e,c,g,1000]}
    tablac = pd.DataFrame(centri)
    print(Fore.GREEN + "__________________________________________________________ \n")

    #linea para exportar datos a excel, y carpeta donde se va a guardar
    tablac.to_excel ('G:/INGENIERIA DE SISTEMAS/SEMESTRE VIII/Linea de Enfasis I/centrifugos.xlsx',index=False)
    
    # Imprimimos la tabla
    print(tablac)
    

    crearTablas('centri')

def absorcion(parametro2):

    g=(parametro2*511.13199046407)/1000		
    c=((parametro2 * 0.0035174111853)*(1925000/0.88))		
    o=c*0.03		
  		
    capex=parametro2*0.0035174111853		
    ft=(capex*1000000)*1.015		
    b=capex*2000000 
    # Creamos la tabla absorción
    absor = {'Energia': ['Microturbina Gas', 'Solar Termica', 'Energia biomasa', 'TR de los chillers de absorcion es:'],
             'Emisiones CO2(TCO2 al mes)':[g,capex,b,""],
             'Capex(dolares megavatios)':[g,capex,b,""],
            'Opex (do-año)': [g,capex,b,1000] 
            }
    tablaab = pd.DataFrame(absor)
    print(Fore.GREEN + "__________________________________________________________ \n")

    #linea para exportar datos a excel, y carpeta donde se va a guardar

    tablaab.to_excel ('G:/INGENIERIA DE SISTEMAS/SEMESTRE VIII/Linea de Enfasis I/absorcion.xlsx',index=False)
    
    # Imprimimos la tabla
    print(tablaab)		
    crearTablas('absor')

def crearTablas(resp):
    if resp == 'centri':
        print(Fore.MAGENTA + " \n Tabla Centrífugos")	
    elif resp == 'absor':
        print(Fore.WHITE + "\n Tabla Absorción")	
init()