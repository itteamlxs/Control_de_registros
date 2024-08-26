Aquí tienes el archivo README.md generado:

```markdown
# Documentación de Macros y Funcionalidad en Excel

Este archivo README.md documenta las macros y la funcionalidad del libro de Excel utilizado para gestionar registros de soporte y procesos en la hoja de trabajo.

## Módulo Registrar

### Descripción General

El **Módulo Registrar** se basa en listas de datos almacenadas en la hoja `Entidades`. Los campos en color rojo claro no deben ser editados, ya que contienen fórmulas que retornan un valor automáticamente. 

- **Encabezado Dinámico**: El título del encabezado es dinámico y se actualiza automáticamente para mostrar el mes actual, basado en una fórmula que toma el valor almacenado en `Entidades`.
- **Botones**: Los botones `REGISTRAR` y `LIMPIAR` están vinculados a código Visual Basic (VB). Para editar estos botones, es necesario ingresar al modo diseño y luego al código. La documentación detallada de la edición de los botones está en el código mismo.

### Campos y Descripción

- **Proceso**: Define el tipo de proceso de mantenimiento a realizar, enlazado a la tabla `Procesos` en la hoja `Entidades`.
- **Continuidad**: Define la recurrencia del proceso (Recurrente, Mensual, Escaso, Semanal), enlazado a la tabla `Recurrencia` en `Entidades`.
- **Librería 1 y 2**: Definen el tipo de librería involucrada en el proceso. No todos los procesos requieren librerías. La tabla `Librerías` también incluye la opción de base de datos en `Entidades`.
- **Cargo**: Cargo de la persona que solicita, enlazado a la tabla `Empleados` en `Entidades`. **No se edita.**
- **Filial**: Filial donde labora la persona que solicita, enlazado a la tabla `Empleados` en `Entidades`. **No se edita.**
- **Fecha de Solicitud**: Fecha en que se solicita el soporte, basada en la función `HOY()` de Excel. **No se edita.**
- **Estado**: Estado de la solicitud de soporte, basado en la tabla `Estados` en `Entidades`.
- **Solicita**: Nombre de quien hace la solicitud de soporte, basado en la tabla `Empleados` en `Entidades`.
- **Cantidad**: Cantidad de procesos que solicitó ese empleado en específico.
- **Total Status proc**: Contiene un total de cada uno de los estados de los procesos realizados, enlazado a la hoja de estadísticas.

### Botones de Almacenamiento y Registro

- **Botón de Almacenamiento**:
  - Cada vez que se almacena un registro, este botón está programado para sumar +1 a la celda F24, generando un número autoincremental que actúa como el número de registro. Este número sirve como un parámetro adicional de seguridad para la impresión.
  - El número autoincremental se puede reiniciar en caso de que se haya perdido un correlativo y sea necesario corregir un registro. Es importante recordar en qué correlativo está actualmente el último registro.

- **Botón de Registro**:
  - El botón graba los rangos de los campos que contienen información importante y restablece a cero las celdas del registro.
  - Este formato está diseñado para ser impreso al momento de entregar el equipo y para poner fecha en el registro al ser devuelto.

### Consulta por Número de Registro

- Se puede buscar cualquier dato importante según el número de registro para verificar su estatus y quién fue la última persona en tenerlo.
- La matriz de consulta permite ver el estatus de un registro físico, si está prestado o devuelto, según el registro físico.

## Código Visual Basic

### Botón de Registrar (CONTROL SOPORTE)

```vb
Private Sub BotonDeRegistrar_Click()

    Dim Hoja1 As Worksheet
    Dim Hoja2 As Worksheet
    Dim ultimaFila As Long
    
    ' Definir las hojas de trabajo
    Set Hoja1 = ThisWorkbook.Sheets("CONTROL SOPORTE")
    Set Hoja2 = ThisWorkbook.Sheets("REGISTROS")
    
    ' Encontrar la última fila en las columnas B a K de Hoja2
    ultimaFila = Hoja2.Cells(Hoja2.Rows.Count, "K").End(xlUp).Row
    
    ' Copiar los valores de C10 a E14 de Hoja1 a la siguiente fila en B a K de Hoja2
    Hoja2.Cells(ultimaFila + 1, "B").Value = Hoja1.Range("C10").Value
    Hoja2.Cells(ultimaFila + 1, "C").Value = Hoja1.Range("C11").Value
    Hoja2.Cells(ultimaFila + 1, "D").Value = Hoja1.Range("C12").Value
    Hoja2.Cells(ultimaFila + 1, "E").Value = Hoja1.Range("C13").Value
    Hoja2.Cells(ultimaFila + 1, "F").Value = Hoja1.Range("C14").Value
    Hoja2.Cells(ultimaFila + 1, "G").Value = Hoja1.Range("E10").Value
    Hoja2.Cells(ultimaFila + 1, "H").Value = Hoja1.Range("E11").Value
    Hoja2.Cells(ultimaFila + 1, "I").Value = Hoja1.Range("E12").Value
    Hoja2.Cells(ultimaFila + 1, "J").Value = Hoja1.Range("E13").Value
    Hoja2.Cells(ultimaFila + 1, "K").Value = Hoja1.Range("E14").Value
    
    ' Limpiar las celdas del formulario en Hoja1
    Hoja1.Range("C10:C13").ClearContents
    Hoja1.Range("E12:E14").ClearContents
    
End Sub
```

### Botón de Registrar (CONTROL ENTRADA SALIDA)

```vb
Private Sub CommandButton1_Click()

    Dim Hoja1 As Worksheet
    Dim Hoja2 As Worksheet
    Dim ultimaFila As Long
    
    ' Definir las hojas de trabajo
    Set Hoja1 = ThisWorkbook.Sheets("CONTROL ENTRADA SALIDA")
    Set Hoja2 = ThisWorkbook.Sheets("REGISTRO ENTRADA_SALIDA")
    
    ' Encontrar la última fila en las columnas B a K de Hoja2
    ultimaFila = Hoja2.Cells(Hoja2.Rows.Count, "K").End(xlUp).Row
    
    ' Copiar los valores de B9 a F17 de Hoja1 a la siguiente fila en B a K de Hoja2
    Hoja2.Cells(ultimaFila + 1, "B").Value = Hoja1.Range("B9").Value
    Hoja2.Cells(ultimaFila + 1, "C").Value = Hoja1.Range("C10").Value
    Hoja2.Cells(ultimaFila + 1, "D").Value = Hoja1.Range("F10").Value
    Hoja2.Cells(ultimaFila + 1, "E").Value = Hoja1.Range("D12").Value
    Hoja2.Cells(ultimaFila + 1, "F").Value = Hoja1.Range("F12").Value
    Hoja2.Cells(ultimaFila + 1, "G").Value = Hoja1.Range("B17").Value
    Hoja2.Cells(ultimaFila + 1, "H").Value = Hoja1.Range("C17").Value
    Hoja2.Cells(ultimaFila + 1, "I").Value = Hoja1.Range("D17").Value
    Hoja2.Cells(ultimaFila + 1, "J").Value = Hoja1.Range("E17").Value
    Hoja2.Cells(ultimaFila + 1, "K").Value = Hoja1.Range("F17").Value
    
    ' Incrementar el valor de la celda F24 en Hoja1
    Hoja1.Range("F24").Value = Hoja1.Range("F24").Value + 1
    
    ' Limpiar las celdas del formulario en Hoja1
    Hoja1.Range("D12").ClearContents
    Hoja1.Range("B17:F23").ClearContents
    
End Sub
```

## Conteo y Gráficos

- **Conteo de recurrencia**: Se genera a partir de la tabla `REGISTROS`, contabilizando cuántos registros se han almacenado de cada proceso.
- **Soporte de filiales**: Se genera a partir de la hoja `REGISTROS`, con el conteo de peticiones por filial.
- **Resultados de status**: Muestran el total de status de cada proceso, categorizados por `Resuelto`, `En Proceso`, `Cancelado`.
- **Gráficos**: Se generan partiendo de las tablas de los conteos por proceso y conteo de soporte por filial.
```
