'''Trabajo Practico Integrador. Organización Empresarial.
ChatBot para alta de proveedores.
(Apellido y nombre de Muriel)
Pisano Di Filippo Mateo Agustín'''


#Importamos la libreria para poder manejar archivos excel
import openpyxl

#Herramienta que nos permite trabajar con la fecha del día. así incorpora automáticamente cuando se le da el alta al proveedor.
from datetime import date 

#variables, archivo excel y la hora correspondiente:
ARCHIVO_EXCEL = "proveedores_murmat.xlsx" 
HOJA = "proveedores"

#función para abrir el excel y cargar a todos los proveedores nuevos a la memoria. Esto nos va a permir verificar cada proveedor para comparar con la nueva entrada
def cargar_proveedores():
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)  #abre el archivo
    hoja = wb[HOJA]   #acceso a la hoja específica
    proveedores = []  #lista vacia para despues llenarla con la nueva info
    for fila in hoja.iter_rows(min_row=2, values_only=True):  #recorre el archivo desde la fila 2 (la fila 1 es el encabezado)
        if fila[0] is not None: #verifica que la fila esté vacia
            #crea un diccionario con los datos del proveedor
            proveedores.append({
                "id": fila[0],
                "nombre": fila[1],
                "cuit": str(fila[2]),
                "rubro": fila[3],
                "contacto": fila[4],
                "email": fila[5],
                "telefono": fila[6],
                "estado": fila[7],
                "fecha_alta": fila[8]
            })
    return proveedores

#funcion para verificar si el proveedor ya está en el sistema:
#Lo hacemos a través del ciut
def proveedor_existe(cuit):
    proveedores = cargar_proveedores()
    for p in proveedores:   #recorre uno por uno cada proveedor
        if p["cuit"] == cuit:   #comparamos el cuit de ingreso con el cuit que está en la lista
            return True  #si ya existe, devuelve True
    return False

#validamos que el cuit sea un número de 11 cifras
def validar_cuit(cuit):
    cuit = cuit.replace("-", "").replace(" ", "")
    if len(cuit) == 11 and cuit.isdigit():
        return cuit
    return None

#validación de que lo ingresado en e-mail sea un e-mal (verificamos que tenga un arroba)
def validar_email(email):
    if "@" in email and "." in email:
        return email
    return None

#validación de que lo ingresado en telefono sea un número de telefono
def validar_telefono(telefono):
    telefono = telefono.replace(" ", "").replace("-", "")
    if telefono.isdigit() and len(telefono) >= 8:
        return telefono
    return None

#Función para agregar al nuevo proveedor al excel
def guardar_proveedor(datos):
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL) #abrimos el excel
    hoja = wb[HOJA] #accedemos a la hora de proveedores
    proveedores = cargar_proveedores() 
    nuevo_id = len(proveedores) + 1  #calcula el ID automáticamente, sumando 1 al último ID
    fecha_hoy = date.today().strftime("%d/%m/%Y")  #agrega la fecha de hoy
    #agregamos una fila nueva
    hoja.append([
        nuevo_id,
        datos["nombre"],
        datos["cuit"],
        datos["rubro"],
        datos["contacto"],
        datos["email"],
        datos["telefono"],
        "activo",
        fecha_hoy
    ])
    #guardamos los cambios
    wb.save(ARCHIVO_EXCEL)

#para 'dar de baja' a un proveedor
def dar_baja_proveedor(cuit):
    #abrimos el excel
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL) 
    #seleccionamos la hoja
    hoja = wb[HOJA] 
    #recorremos las filas
    for fila in hoja.iter_rows(min_row = 2):
        if str(fila[2].value) == cuit:
            #damos los datos al usuario para que lo visualice a la hora de la confirmación.
            print("\n---Proveedor encontrado.") 
            print(f"  Nombre:   {fila[1].value}")
            print(f"  CUIT:     {fila[2].value}")
            print(f"  Rubro:    {fila[3].value}")
            print(f"  Contacto: {fila[4].value}")
            print("---------------------------")
            
            #hacemos una confirmación breve por si el usuario se equivocó
            while True:
                confirmar = input("¿Confirmás la baja? (s/n): ").strip().lower()
                # en caso de 's', se elimina al proveedor completamente
                if confirmar == "s":
                    hoja.delete_rows(fila[0].row)
                    wb.save(ARCHIVO_EXCEL)
                    print("Proveedor eliminado del sistema exitosamente.")
                    return
                elif confirmar == "n":
                    print("Operación cancelada.")
                    return
                else:
                    print("Ingresá 's' para confirmar o 'n' para cancelar.")
    #en el caso de que el cuit sea incorrecto
    print("No se encontró ningún proveedor con ese CUIT.") 


#para cambiar de estado al proveedor (de activo a inactivo y viceversa)
def cambiar_estados_proveedor(cuit):
    #abrimos el excel
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL) 
    #seleccionamos la hoja
    hoja = wb[HOJA] 
    #recorremos las filas
    for fila in hoja.iter_rows(min_row=2):
        if str(fila[2].value) == cuit:
            estado_actual = fila[7].value
            if estado_actual == "activo":
                nuevo_estado = "inactivo"
            else:
                nuevo_estado = "activo" 
            
            #mostramos en pantalla al proveedor.
            print("\n---Proveedor encontrado.")
            print(f"  Nombre:   {fila[1].value}")
            print(f"  CUIT:     {fila[2].value}")
            print(f"  Rubro:    {fila[3].value}")
            print(f"  Contacto: {fila[4].value}")
            print(f"  Estado:   {fila[7].value}")
            print("---------------------------")
            
            #confirmación antes de reactivar
            while True:
                confirmar = input("¿Confirmás la reactivación? (s/n): ").strip().lower()
                if confirmar == "s":
                    #cambiamos el estado del proveedor al estado contrario.
                    fila[7].value = nuevo_estado  
                    wb.save(ARCHIVO_EXCEL)
                    print(f"\nEstado cambiado a '{nuevo_estado}' exitosamente.")
                    return
                elif confirmar == "n":
                    print("Operación cancelada.")
                    return
                else:
                    print("Ingresá 's' para confirmar o 'n' para cancelar.")
    print("No se encontró ningún proveedor con ese CUIT.")

#Para dar de alta a los proveedores:
def alta_proveedor():
    #abrimos el diccionario
    datos = {}
 #nombre
    while True:
        nombre = input("Nombre o razón social del proveedor: ").strip()
        if nombre.lower() == "salir":
            print("Operación cancelada.")
            return
        if len(nombre) == 0:
            print("Debe ingresar algo. Intentá de nuevo.")
            continue
        datos["nombre"] = nombre
        break
#cuit
    while True:
        cuit_input = input("CUIT de 11 dígitos, con o sin guiones: ").strip()
        if cuit_input.lower() == "salir":
            print("Operación cancelada.")
            return
        cuit = validar_cuit(cuit_input)
        if not cuit:
            print("CUIT inválido. Debe tener 11 dígitos numéricos.")
            continue
        if proveedor_existe(cuit):
            print("Este proveedor ya existe en el sistema. Operación cancelada.")
            return
        datos["cuit"] = cuit
        break
#rubro
    while True:
        rubro = input("Rubro: ").strip()
        if rubro.lower() == "salir":
            print("Operación cancelada.")
            return
        if len(rubro) == 0:
            print("Debe ingresar algo. Intentá de nuevo.")
            continue
        datos["rubro"] = rubro
        break
#contacto
    while True:
        contacto = input("Nombre de la persona de contacto: ").strip()
        if contacto.lower() == "salir":
            print("Operación cancelada.")
            return
        if len(contacto) == 0:
            print("Debe ingresar algo. Intentá de nuevo.")
            continue
        datos["contacto"] = contacto
        break
#email
    while True:
        email = input("Email de contacto: ").strip()
        if email.lower() == "salir":
            print("Operación cancelada.")
            return
        if not validar_email(email):
            print("Email inválido. Debe contener @ y un dominio.")
            continue
        datos["email"] = email
        break
#telefono
    while True:
        telefono = input("Teléfono de contacto: ").strip()
        if telefono.lower() == "salir":
            print("Operación cancelada.")
            return
        if not validar_telefono(telefono):
            print("Teléfono inválido. Solo números, mínimo 8 dígitos.")
            continue
        datos["telefono"] = telefono
        break
#datos para que el usuario chequee que todo esté bien
    print("\n--- Resumen del nuevo proveedor ---")
    print(f"  Nombre:   {datos['nombre']}")
    print(f"  CUIT:     {datos['cuit']}")
    print(f"  Rubro:    {datos['rubro']}")
    print(f"  Contacto: {datos['contacto']}")
    print(f"  Email:    {datos['email']}")
    print(f"  Teléfono: {datos['telefono']}")
    print("-----------------------------------")
#pequeña confirmación por si algo no está bien
    while True:
        confirmar = input("¿Confirmás el alta? (s/n): ").strip().lower()
        if confirmar == "s":
            guardar_proveedor(datos)
            print("Proveedor dado de alta exitosamente en el sistema.")
            return
        elif confirmar == "n":
            print("Operación cancelada.")
            return
        else:
            print("Ingresá 's' para confirmar o 'n' para cancelar.")

#mostrar lista completa de proveedores
def mostrar_proveedores():
    proveedores = cargar_proveedores() 
    if len(proveedores) == 0:
        print("No hay proveedores cargados en el sistema.")
        return
    print('''---Lista de Proveedores---''') 
    for p in proveedores: #recorremos el archivo con un for
        print(f"""
ID:           {p['id']}
Nombre:       {p['nombre']}
CUIT:         {p['cuit']}
Rubro:        {p['rubro']}
Contacto:     {p['contacto']}
Email:        {p['email']}
Teléfono:     {p['telefono']}
Estado:       {p['estado']}
Fecha de alta:{p['fecha_alta']}
---------------------------""")


#Funcion para el menú
def menu():
    
    #titulo
    print()
    print("-" * 45)
    print("MURMAT S.A. — Sistema de Alta de Proveedores")
    print("-" * 45)
    print("Escribí 'salir' en cualquier momento para cancelar.\n")

    #Preguntamos que acción quiere tomar (alta o baja)
    while True:
        print("\n¿Qué deseas hacer?")
        print("""
1. Dar de alta un proveedor
2. Dar de baja un proveedor
3. Cambiar estado de un proveedor (Activo/Inactivo)
4. Ver lista completa de proveedores
5. Salir
""")
        opcion = input('''Elegí una opción (1/2/3/4/5):
''').strip()
        #salir
        if opcion == '5':
            print('Programa finalizado.')
            break
        #alta
        elif opcion == "1":
            alta_proveedor()
        #baja
        elif opcion == "2":
            while True:
                cuit_input = input("Ingresá el CUIT del proveedor a dar de baja: ").strip()
                if cuit_input.lower() == "salir":
                    print("Operación cancelada.")
                    break
                cuit = validar_cuit(cuit_input)
                if not cuit:
                    print("CUIT inválido. Debe tener 11 dígitos numéricos.")
                    continue
                dar_baja_proveedor(cuit)
                break
        
        #reactivación de un proveedor dado de baja
        elif opcion == "3":
            while True:
                cuit_input = input("Ingresá el CUIT del proveedor a cambiar su estado: ").strip()
                if cuit_input.lower() == "salir":
                    print("Operación cancelada.")
                    break
                cuit = validar_cuit(cuit_input)
                if not cuit:
                    print("CUIT inválido. Debe tener 11 dígitos numéricos.")
                    continue
                cambiar_estados_proveedor(cuit)
                break
        
        #mostrar lista proveedores
        elif opcion == '4':
            mostrar_proveedores()
        else:
            print("Opción inválida. Ingresá 1,2,3,4,5.\n")

#invocamos a menú
menu()