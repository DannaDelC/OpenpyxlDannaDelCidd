import openpyxl
#Danna Lucrecia Del Cid López

def crear_informe_gastos(nombre_archivo):
    
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Gastos"
    
    hoja['A1'] = "Fecha"
    hoja['B1'] = "Descripción"
    hoja['C1'] = "Monto"
    
    
    total_gastos = 0
    gasto_mas_caro = None
    gasto_mas_barato = None
    max_monto = float('-inf')
    min_monto = float('inf')
    num_gastos = 0
    
    
    while True:
        fecha = input("Ingrese la fecha del gasto (o 'fin' para terminar): ")
        if fecha.lower() == 'fin':
            break
        
        descripcion = input("Ingrese la descripción del gasto: ")
        try:
            monto = float(input("Ingrese el monto del gasto: "))
        except ValueError:
            print("Error: Ingrese un monto válido.")
            continue
        
        
        num_gastos += 1
        hoja.append([fecha, descripcion, monto])
        
        
        total_gastos += monto
        
    
        if monto > max_monto:
            max_monto = monto
            gasto_mas_caro = (fecha, descripcion)
        if monto < min_monto:
            min_monto = monto
            gasto_mas_barato = (fecha, descripcion)
    
    
    libro.save(nombre_archivo)
    
    
    print("\nResumen de Gastos:")
    print(f"Número total de gastos: {num_gastos}")
    if gasto_mas_caro:
        print(f"Gasto más caro - Fecha: {gasto_mas_caro[0]}, Descripción: {gasto_mas_caro[1]}, Monto: {max_monto}")
    if gasto_mas_barato:
        print(f"Gasto más barato - Fecha: {gasto_mas_barato[0]}, Descripción: {gasto_mas_barato[1]}, Monto: {min_monto}")
    print(f"Monto total de gastos: {total_gastos}")
    print(f"\nSe ha guardado el informe de gastos en el archivo: '{nombre_archivo}'")


crear_informe_gastos("informe_gastos.xlsx")
