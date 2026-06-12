'''Trabajo Practico Integrador. Organización Empresarial.
ChatBot para alta de proveedores - VERSIÓN CON MÁQUINA DE ESTADOS (FSM)
Heinzle Muriel Amancay. Comisión 8
Pisano Di Filippo Mateo Agustín. Comisión 8'''

import openpyxl
from datetime import date

ARCHIVO_EXCEL = "proveedores_murmat.xlsx"
HOJA = "proveedores"

# =========================================================
#  FUNCIONES DE DATOS Y VALIDACIÓN (sin cambios respecto al original)
# =========================================================

def cargar_proveedores():
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb[HOJA]
    proveedores = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None:
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


def proveedor_existe(cuit):
    proveedores = cargar_proveedores()
    for p in proveedores:
        if p["cuit"] == cuit:
            return True
    return False


def validar_cuit(cuit):
    cuit = cuit.replace("-", "").replace(" ", "")
    if len(cuit) == 11 and cuit.isdigit():
        return cuit
    return None


def validar_email(email):
    if "@" in email and "." in email:
        return email
    return None


def validar_telefono(telefono):
    telefono = telefono.replace(" ", "").replace("-", "")
    if telefono.isdigit() and len(telefono) >= 8:
        return telefono
    return None


def guardar_proveedor(datos):
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb[HOJA]
    proveedores = cargar_proveedores()
    nuevo_id = len(proveedores) + 1
    fecha_hoy = date.today().strftime("%d/%m/%Y")
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
    wb.save(ARCHIVO_EXCEL)


def buscar_por_nombre(nombre_busqueda):
    proveedores = cargar_proveedores()
    for i in proveedores:
        if nombre_busqueda.lower() in i['nombre'].lower():
            return i['cuit']
    print('\nNo se encontró ningún proveedor con ese nombre.')
    return None


def mostrar_proveedores():
    proveedores = cargar_proveedores()
    if len(proveedores) == 0:
        print('No hay proveedores cargados en el sistema.')
        return
    print('''---Lista de Proveedores---''')
    for p in proveedores:
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


def normalizar_cuit_celda(valor):
    """Normaliza el valor leído de una celda para comparar CUITs."""
    return str(valor).strip().replace("-", "").replace(" ", "")


# =========================================================
#  MÁQUINA DE ESTADOS
# =========================================================
#
# 'contexto' guarda TODO lo que el bot necesita "recordar":
#   - estado_actual: en qué paso estamos
#   - datos: lo que se va completando (formulario de alta, etc.)
#   - cuit_objetivo: el proveedor sobre el que se está operando
#
# Cada función "manejar_<estado>" recibe el contexto, pide UN dato
# (o ejecuta UNA acción), y devuelve el nombre del PRÓXIMO estado.
# El bucle principal solo se encarga de llamar a la función correcta
# según contexto["estado_actual"].
# =========================================================

def contexto_inicial():
    return {
        "estado_actual": "menu",
        "datos": {},          # datos parciales del alta en curso
        "cuit_objetivo": None # cuit sobre el que se opera (baja/estado/modificar)
    }


# ---------------------------------------------------------
# Estado: MENU PRINCIPAL
# ---------------------------------------------------------
def manejar_menu(ctx):
    print("\n¿Qué deseas hacer?")
    print("""
1. Dar de alta un proveedor
2. Dar de baja un proveedor
3. Cambiar estado de un proveedor (Activo/Inactivo)
4. Ver lista completa de proveedores
5. Buscar proveedor por nombre/CUIT
6. Modificar datos del proveedor
7. Salir
""")
    opcion = input("Elegí una opción:\n").strip()

    if opcion == "1":
        ctx["datos"] = {}
        return "alta_nombre"
    elif opcion == "2":
        return "baja_buscar"
    elif opcion == "3":
        return "estado_buscar"
    elif opcion == "4":
        mostrar_proveedores()
        return "menu"
    elif opcion == "5":
        return "ver_buscar"
    elif opcion == "6":
        return "modificar_buscar"
    elif opcion == "7":
        print("Programa finalizado.")
        return "fin"
    else:
        print("\nOpción inválida.\n")
        return "menu"


# ---------------------------------------------------------
# ALTA DE PROVEEDOR — cada campo es un estado
# ---------------------------------------------------------
def manejar_alta_nombre(ctx):
    nombre = input("Nombre o razón social del proveedor: ").strip().title()
    if nombre.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    if len(nombre) == 0:
        print("Debe ingresar algo. Intentá de nuevo.")
        return "alta_nombre"
    ctx["datos"]["nombre"] = nombre
    return "alta_cuit"


def manejar_alta_cuit(ctx):
    cuit_input = input("CUIT de 11 dígitos, con o sin guiones: ").strip()
    if cuit_input.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    cuit = validar_cuit(cuit_input)
    if not cuit:
        print("CUIT inválido. Debe tener 11 dígitos numéricos.")
        return "alta_cuit"
    if proveedor_existe(cuit):
        print("Este proveedor ya existe en el sistema. Operación cancelada.")
        return "menu"
    ctx["datos"]["cuit"] = cuit
    return "alta_rubro"


def manejar_alta_rubro(ctx):
    rubro = input("Rubro: ").strip().title()
    if rubro.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    if len(rubro) == 0:
        print("Debe ingresar algo. Intentá de nuevo.")
        return "alta_rubro"
    ctx["datos"]["rubro"] = rubro
    return "alta_contacto"


def manejar_alta_contacto(ctx):
    contacto = input("Nombre de la persona de contacto: ").strip().title()
    if contacto.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    if len(contacto) == 0:
        print("Debe ingresar algo. Intentá de nuevo.")
        return "alta_contacto"
    ctx["datos"]["contacto"] = contacto
    return "alta_email"


def manejar_alta_email(ctx):
    email = input("Email de contacto: ").strip()
    if email.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    if not validar_email(email):
        print("Email inválido. Debe contener @ y un dominio.")
        return "alta_email"
    ctx["datos"]["email"] = email
    return "alta_telefono"


def manejar_alta_telefono(ctx):
    telefono = input("Teléfono de contacto: ").strip()
    if telefono.lower() == "salir":
        print("Operación cancelada.")
        return "menu"
    if not validar_telefono(telefono):
        print("Teléfono inválido. Solo números, mínimo 8 dígitos.")
        return "alta_telefono"
    ctx["datos"]["telefono"] = telefono
    return "alta_confirmar"


def manejar_alta_confirmar(ctx):
    datos = ctx["datos"]
    print("\n--- Resumen del nuevo proveedor ---")
    print(f"  Nombre:   {datos['nombre']}")
    print(f"  CUIT:     {datos['cuit']}")
    print(f"  Rubro:    {datos['rubro']}")
    print(f"  Contacto: {datos['contacto']}")
    print(f"  Email:    {datos['email']}")
    print(f"  Teléfono: {datos['telefono']}")
    print("-----------------------------------")

    confirmar = input("¿Confirmás el alta? (s/n): ").strip().lower()
    if confirmar == "s":
        guardar_proveedor(datos)
        print("Proveedor dado de alta exitosamente en el sistema.")
        return "menu"
    elif confirmar == "n":
        print("Operación cancelada.")
        return "menu"
    else:
        print("Ingresá 's' para confirmar o 'n' para cancelar.")
        return "alta_confirmar"


# ---------------------------------------------------------
# BAJA DE PROVEEDOR
# ---------------------------------------------------------
def manejar_baja_buscar(ctx):
    cuit_nombre = input('\nIngresá el CUIT o nombre del proveedor a dar de baja: ').strip()
    if cuit_nombre.lower() == 'salir':
        print('\nOperación cancelada.')
        return "menu"

    if cuit_nombre.replace('-', '').replace(' ', '').isdigit():
        cuit = validar_cuit(cuit_nombre)
        if not cuit:
            print("CUIT inválido. Debe tener 11 dígitos numéricos.")
            return "baja_buscar"
    else:
        cuit = buscar_por_nombre(cuit_nombre)
        if not cuit:
            return "menu"

    ctx["cuit_objetivo"] = cuit
    return "baja_confirmar"


def manejar_baja_confirmar(ctx):
    cuit = ctx["cuit_objetivo"]
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb[HOJA]

    for fila in hoja.iter_rows(min_row=2):
        if normalizar_cuit_celda(fila[2].value) == cuit:
            print("\n---Proveedor encontrado.")
            print(f"  Nombre:   {fila[1].value}")
            print(f"  CUIT:     {fila[2].value}")
            print(f"  Rubro:    {fila[3].value}")
            print(f"  Contacto: {fila[4].value}")
            print("---------------------------")

            confirmar = input("¿Confirmás la baja? (s/n): ").strip().lower()
            if confirmar == "s":
                hoja.delete_rows(fila[0].row)
                wb.save(ARCHIVO_EXCEL)
                print("Proveedor eliminado del sistema exitosamente.")
                return "menu"
            elif confirmar == "n":
                print("Operación cancelada.")
                return "menu"
            else:
                print("Ingresá 's' para confirmar o 'n' para cancelar.")
                return "baja_confirmar"

    print("No se encontró ningún proveedor con ese CUIT.")
    return "menu"


# ---------------------------------------------------------
# CAMBIAR ESTADO (activo/inactivo)
# ---------------------------------------------------------
def manejar_estado_buscar(ctx):
    cuit_nombre = input('\nIngresá el CUIT o el nombre del proveedor a cambiar su estado: ').strip()
    if cuit_nombre.lower() == 'salir':
        print('\nOperación cancelada.')
        return "menu"

    if cuit_nombre.replace('-', '').replace(' ', '').isdigit():
        cuit = validar_cuit(cuit_nombre)
        if not cuit:
            print("CUIT inválido. Debe tener 11 dígitos numéricos.")
            return "estado_buscar"
    else:
        cuit = buscar_por_nombre(cuit_nombre)
        if not cuit:
            return "menu"

    ctx["cuit_objetivo"] = cuit
    return "estado_confirmar"


def manejar_estado_confirmar(ctx):
    cuit = ctx["cuit_objetivo"]
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb[HOJA]

    for fila in hoja.iter_rows(min_row=2):
        if normalizar_cuit_celda(fila[2].value) == cuit:
            estado_actual = fila[7].value
            nuevo_estado = "inactivo" if estado_actual == "activo" else "activo"

            print("\n---Proveedor encontrado.")
            print(f"  Nombre:   {fila[1].value}")
            print(f"  CUIT:     {fila[2].value}")
            print(f"  Rubro:    {fila[3].value}")
            print(f"  Contacto: {fila[4].value}")
            print(f"  Estado:   {fila[7].value}")
            print("---------------------------")

            confirmar = input("¿Confirmás la acción? (s/n): ").strip().lower()
            if confirmar == "s":
                fila[7].value = nuevo_estado
                wb.save(ARCHIVO_EXCEL)
                print(f"\nEstado cambiado a '{nuevo_estado}' exitosamente.")
                return "menu"
            elif confirmar == "n":
                print("Operación cancelada.")
                return "menu"
            else:
                print("Ingresá 's' para confirmar o 'n' para cancelar.")
                return "estado_confirmar"

    print("No se encontró ningún proveedor con ese CUIT.")
    return "menu"


# ---------------------------------------------------------
# VER PROVEEDOR
# ---------------------------------------------------------
def manejar_ver_buscar(ctx):
    cuit_nombre = input('\nIngresá el CUIT o nombre del proveedor a buscar: ').strip()
    if cuit_nombre.lower() == 'salir':
        print('\nOperación cancelada.')
        return "menu"

    if cuit_nombre.replace('-', '').replace(' ', '').isdigit():
        cuit = validar_cuit(cuit_nombre)
        if not cuit:
            print("CUIT inválido. Debe tener 11 dígitos numéricos.")
            return "ver_buscar"
    else:
        cuit = buscar_por_nombre(cuit_nombre)
        if not cuit:
            return "menu"

    proveedores = cargar_proveedores()
    for i in proveedores:
        if i["cuit"] == cuit:
            print(f"""
ID:           {i['id']}
Nombre:       {i['nombre']}
CUIT:         {i['cuit']}
Rubro:        {i['rubro']}
Contacto:     {i['contacto']}
Email:        {i['email']}
Teléfono:     {i['telefono']}
Estado:       {i['estado']}
Fecha de alta:{i['fecha_alta']}
""")
            return "menu"

    print('\nNo se encontró el proveedor.')
    return "menu"


# ---------------------------------------------------------
# MODIFICAR PROVEEDOR
# ---------------------------------------------------------
def manejar_modificar_buscar(ctx):
    cuit_nombre = input('\nIngresá el CUIT o nombre del proveedor a modificar: ').strip()
    if cuit_nombre.lower() == 'salir':
        print('\nOperación cancelada.')
        return "menu"

    if cuit_nombre.replace('-', '').replace(' ', '').isdigit():
        cuit = validar_cuit(cuit_nombre)
        if not cuit:
            print("CUIT inválido. Debe tener 11 dígitos numéricos.")
            return "modificar_buscar"
    else:
        cuit = buscar_por_nombre(cuit_nombre)
        if not cuit:
            return "menu"

    ctx["cuit_objetivo"] = cuit
    return "modificar_editar"


def manejar_modificar_editar(ctx):
    cuit = ctx["cuit_objetivo"]
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb[HOJA]

    for fila in hoja.iter_rows(min_row=2):
        if normalizar_cuit_celda(fila[2].value) == cuit:
            print('\n---Proveedor encontrado.')
            print(f'  Nombre:   {fila[1].value}')
            print(f'  CUIT:     {fila[2].value}')
            print(f'  Rubro:    {fila[3].value}')
            print(f'  Contacto: {fila[4].value}')
            print(f'  Email:    {fila[5].value}')
            print(f'  Teléfono: {fila[6].value}')

            print('\nEn caso de no querer modificar el campo, dejar vacio y presionar enter.')
            print('Escribí "salir" en cualquier momento para cancelar la modificación.\n')

            nombre = input(f'Nuevo nombre [{fila[1].value}]: ').strip()
            if nombre.lower() == "salir":
                print('\nOperación cancelada.')
                return "menu"
            if nombre:
                fila[1].value = nombre.title()

            while True:
                cuit_nuevo = input(f'Nuevo CUIT [{fila[2].value}]: ').strip()
                if cuit_nuevo.lower() == "salir":
                    print('\nOperación cancelada.')
                    return "menu"
                if not cuit_nuevo:
                    break
                cuit_validado = validar_cuit(cuit_nuevo)
                if not cuit_validado:
                    print("CUIT inválido. Debe tener 11 dígitos numéricos.")
                    continue
                if cuit_validado != cuit and proveedor_existe(cuit_validado):
                    print("Ya existe otro proveedor con ese CUIT.")
                    continue
                fila[2].value = cuit_validado
                break

            rubro = input(f"Nuevo rubro [{fila[3].value}]: ").strip()
            if rubro.lower() == "salir":
                print('\nOperación cancelada.')
                return "menu"
            if rubro:
                fila[3].value = rubro.title()

            contacto = input(f"Nuevo contacto [{fila[4].value}]: ").strip()
            if contacto.lower() == "salir":
                print('\nOperación cancelada.')
                return "menu"
            if contacto:
                fila[4].value = contacto.title()

            while True:
                email = input(f"Nuevo email [{fila[5].value}]: ").strip()
                if email.lower() == "salir":
                    print('\nOperación cancelada.')
                    return "menu"
                if not email:
                    break
                if validar_email(email):
                    fila[5].value = email
                    break
                print("Email inválido. Debe contener @ y un dominio.")

            while True:
                telefono = input(f"Nuevo teléfono [{fila[6].value}]: ").strip()
                if telefono.lower() == "salir":
                    print('\nOperación cancelada.')
                    return "menu"
                if not telefono:
                    break
                telefono_validado = validar_telefono(telefono)
                if telefono_validado:
                    fila[6].value = telefono_validado
                    break
                print("Teléfono inválido. Solo números, mínimo 8 dígitos.")

            confirmar = input("\n¿Confirmás los cambios? (s/n): ").strip().lower()
            if confirmar == "s":
                wb.save(ARCHIVO_EXCEL)
                print("\nProveedor actualizado exitosamente.")
                return "menu"
            elif confirmar == "n":
                print("\nOperación cancelada.")
                return "menu"
            else:
                print('\nIngresá "s" o "n".')
                # Nota: si releemos el archivo, los cambios sin guardar
                # de esta pasada se pierden y se vuelve a pedir todo.
                return "modificar_editar"

    print("\nNo se encontró ningún proveedor con ese CUIT.")
    return "menu"


# =========================================================
#  TABLA DE TRANSICIONES (despachador)
# =========================================================
# Diccionario que mapea cada nombre de estado a la función
# que lo maneja. Esto reemplaza toda la cascada de if/elif
# del menú original.
MANEJADORES = {
    "menu": manejar_menu,

    "alta_nombre": manejar_alta_nombre,
    "alta_cuit": manejar_alta_cuit,
    "alta_rubro": manejar_alta_rubro,
    "alta_contacto": manejar_alta_contacto,
    "alta_email": manejar_alta_email,
    "alta_telefono": manejar_alta_telefono,
    "alta_confirmar": manejar_alta_confirmar,

    "baja_buscar": manejar_baja_buscar,
    "baja_confirmar": manejar_baja_confirmar,

    "estado_buscar": manejar_estado_buscar,
    "estado_confirmar": manejar_estado_confirmar,

    "ver_buscar": manejar_ver_buscar,

    "modificar_buscar": manejar_modificar_buscar,
    "modificar_editar": manejar_modificar_editar,
}


# =========================================================
#  BUCLE PRINCIPAL
# =========================================================
def main():
    print()
    print("-" * 45)
    print("MURMAT S.A. — Sistema de Alta de Proveedores")
    print("-" * 45)
    print("Escribí 'salir' en cualquier momento para cancelar.\n")

    ctx = contexto_inicial()

    # El bucle no sabe NADA sobre alta/baja/modificar.
    # Solo mira ctx["estado_actual"], busca la función correspondiente
    # en MANEJADORES, la ejecuta, y guarda el estado que ésta devuelve.
    while ctx["estado_actual"] != "fin":
        manejador = MANEJADORES[ctx["estado_actual"]]
        siguiente_estado = manejador(ctx)
        ctx["estado_actual"] = siguiente_estado


if __name__ == "__main__":
    main()